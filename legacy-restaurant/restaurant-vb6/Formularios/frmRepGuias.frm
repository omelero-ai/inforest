VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepGuias 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guias "
   ClientHeight    =   3675
   ClientLeft      =   2895
   ClientTop       =   2190
   ClientWidth     =   8310
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepGuias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   8310
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
      Left            =   5415
      Picture         =   "frmRepGuias.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3030
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
      Left            =   6870
      Picture         =   "frmRepGuias.frx":10EC
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3030
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
      Left            =   3960
      Picture         =   "frmRepGuias.frx":11DE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3030
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
      Left            =   2505
      Picture         =   "frmRepGuias.frx":1710
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3030
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   6960
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Height          =   2955
      Left            =   0
      TabIndex        =   10
      Top             =   15
      Width           =   8325
      Begin VB.CheckBox chkEstado 
         Caption         =   "Todos los Estados"
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
         Left            =   5730
         TabIndex        =   17
         Top             =   1620
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.TextBox txtTransportista 
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
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2040
         Width           =   2775
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
         Height          =   375
         Index           =   1
         Left            =   4830
         Picture         =   "frmRepGuias.frx":1C42
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1980
         Width           =   765
      End
      Begin VB.CheckBox chkMotivo 
         Caption         =   "Todos los Motivos"
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
         Left            =   5730
         TabIndex        =   5
         Top             =   2445
         Value           =   1  'Checked
         Width           =   2325
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
         Height          =   375
         Index           =   0
         Left            =   4845
         Picture         =   "frmRepGuias.frx":1D44
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1170
         Width           =   765
      End
      Begin VB.TextBox txtCliente 
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
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CheckBox chkTransportista 
         Caption         =   "Todos los Transportistas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5730
         TabIndex        =   3
         Top             =   2055
         Value           =   1  'Checked
         Width           =   2490
      End
      Begin VB.CheckBox chkCliente 
         Caption         =   "Todos los Clientes"
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
         Left            =   5760
         TabIndex        =   2
         Top             =   1215
         Value           =   1  'Checked
         Width           =   2340
      End
      Begin MSDataListLib.DataCombo cboMotivo 
         Height          =   315
         Left            =   1995
         TabIndex        =   4
         Top             =   2445
         Width           =   3615
         _ExtentX        =   6376
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
      Begin MSDataListLib.DataCombo cboEstado 
         Height          =   315
         Left            =   1995
         TabIndex        =   18
         Top             =   1605
         Width           =   3615
         _ExtentX        =   6376
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
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   315
         Left            =   1995
         TabIndex        =   20
         Top             =   765
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
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
         Format          =   51642369
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   315
         Left            =   1995
         TabIndex        =   22
         Top             =   345
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
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
         Format          =   51642369
         CurrentDate     =   37539
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial :"
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
         Left            =   690
         TabIndex        =   21
         Top             =   405
         Width           =   1230
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado :"
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
         Left            =   1200
         TabIndex        =   19
         Top             =   1665
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo :"
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
         Index           =   7
         Left            =   1215
         TabIndex        =   14
         Top             =   2490
         Width           =   705
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transportista :"
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
         Index           =   5
         Left            =   675
         TabIndex        =   13
         Top             =   2070
         Width           =   1245
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente :"
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
         Index           =   10
         Left            =   1200
         TabIndex        =   12
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   930
         TabIndex        =   11
         Top             =   825
         Width           =   990
      End
   End
End
Attribute VB_Name = "frmRepGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Reporte         As New dsrGuia
Dim RsReporte       As Recordset
Dim RsMotivo        As Recordset
Dim RsEstado        As Recordset

Dim sCriterio       As String
Dim sCliente        As String
Dim sTransportista  As String

Private Sub cmdBusca_Click(Index As Integer)
    If Index = 0 Then
        'Cliente Solo los que tienen numero de cliente, los demas no se muestran
        'no es lo mismo pues los valores q retorna el BUSCA se asignan a diferentes variables, por eso estan dentro del IF_ELSE_ENDIF
        Isql = "SELECT DISTINCT dbo.vCliente.Codigo AS Codigo, dbo.vCliente.Descripcion " & _
                "FROM      dbo.MDOCUMENTO LEFT OUTER JOIN " & _
                "dbo.vCliente ON dbo.MDOCUMENTO.tCodigoCliente = dbo.vCliente.Codigo " & _
                "WHERE     (ISNULL(dbo.MDOCUMENTO.tCodigoCliente, '') <> '') AND (dbo.MDOCUMENTO.tEstadoDocumento <> '04') OR " & _
                "(dbo.MDOCUMENTO.tEstadoDocumento <> '04') AND (LEN(LTRIM(dbo.MDOCUMENTO.tCodigoCliente)) <> 0)"
        frmBusca.cboCriterio.Enabled = True
        Call ConfGrilla(2, frmBusca.grdGrilla, "Razón Comercial", 2, "Descripcion", 7300, 0, 0, "", _
                                               "CodigoCliente", 2, "Codigo", 2600, 0, 0, "")
        frmBusca.grdGrilla.Columns(1).Visible = False
        frmBusca.Show vbModal
        If Not wEnter Then
           Exit Sub
        End If
        sCliente = sCodigo
        txtCliente = sDescrip
    ElseIf Index = 1 Then
        'Transportista Aca si se muestran todos los clientes
        Isql = "Select Codigo, Descripcion from vCLIENTE order by Descripcion"
        Call ConfGrilla(2, frmBusca.grdGrilla, "Razón Comercial", 2, "Descripcion", 7000, 0, 0, "", _
                                               "CodigoCliente", 2, "Codigo", 2600, 0, 0, "")
        frmBusca.grdGrilla.Columns(1).Visible = False
        frmBusca.Show vbModal
        If Not wEnter Then
           Exit Sub
        End If
        sTransportista = sCodigo
        txtTransportista = sDescrip
    End If
    
End Sub

Private Sub chkCliente_Click()
'   CambioEstado_Combo txtCliente, chkCliente.Value
   If chkCliente.Value = 1 Then
      cmdBusca(0).Enabled = False
      txtCliente = ""
      txtCliente.Locked = True
   Else
      cmdBusca(0).Enabled = True
   End If
End Sub

Private Sub chkMotivo_Click()
    If chkMotivo.Value = 1 Then
      cboMotivo.Enabled = False
      cboMotivo.Text = ""
   Else
      cboMotivo.Enabled = True
   End If
End Sub

Private Sub chkTransportista_Click()
'   CambioEstado_Combo txtTransportista, chkTransportista.Value
   If chkTransportista.Value = 1 Then
      cmdBusca(1).Enabled = False
      txtTransportista = ""
      txtTransportista.Locked = True
   Else
      cmdBusca(1).Enabled = True
      
   End If
End Sub

Private Sub chkEstado_Click()
 'CambioEstado_Combo cboEstado, chkEstado.Value
   If chkEstado.Value = 1 Then
      cboEstado.Enabled = False
      cboEstado.Text = ""
   Else
      cboEstado.Enabled = True
   End If
 
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   
   sCriterio = ""
   
   If dtpFecIni.Value > dtpFecFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      Exit Sub
      dtpFecFin.SetFocus
   End If
   sCriterio = sCriterio & "fFecha>='" & Format(dtpFecIni.Value, "YYYY-MM-DD") & " 00:00' AND fFecha<='" & Format(dtpFecFin.Value, "YYYY-MM-DD") & " 23:59'"
   'Cliente
   If chkCliente.Value = 0 Then
      If sCliente = "" Then
         MsgBox "Debe escoger el Cliente", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " AND tDestinatario ='" & sCliente & "'"
   End If
   'Transportista
    If chkTransportista.Value = 0 Then
        If sTransportista = "" Then
            MsgBox "Debe escoger un Transportista", vbCritical, sMensaje
            Exit Sub
        End If
        sCriterio = sCriterio & " AND tTransportista ='" & sTransportista & "'"
    End If
    'Motivo
    If chkMotivo.Value = 0 Then
        If cboMotivo.Text = "" Then
            MsgBox "Debe escoger un Motivo", vbCritical, sMensaje
            Exit Sub
        End If
        sCriterio = sCriterio & " AND tMotivo ='" & cboMotivo.BoundText & "'"
    End If
    'Estado
    If chkEstado.Value = 0 Then
        If cboEstado.Text = "" Then
            MsgBox "Debe escoger un Estado", vbCritical, sMensaje
            Exit Sub
        End If
        sCriterio = sCriterio & " AND tEstado ='" & cboEstado.BoundText & "'"
    End If

    Select Case Index
        Case Is = 0 ' Preview
            Genera
            If RsReporte.EOF = True Then
                Screen.MousePointer = vbDefault
                MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                Exit Sub
            End If
                Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                Reporte.PaperOrientation = crPortrait
               
                frmEmite.CRViewer.ViewReport
                frmEmite.Show vbModal
        Case Is = 1 ' Imprimir
            Genera
            Screen.MousePointer = vbDefault
            If RsReporte.EOF = True Then
                MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                Exit Sub
            End If
            Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
            Reporte.PaperOrientation = crPortrait
            Reporte.PrintOut
        Case Is = 2 ' Salir
               Unload Me
        Case Is = 3 ' Exportar
            Genera
            Screen.MousePointer = vbDefault
            If RsReporte.EOF = True Then
               MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
               Exit Sub
            End If
                           
            Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
            Reporte.PaperOrientation = crPortrait
            Reporte.ExportOptions.FormatType = 21
            Reporte.ExportOptions.DestinationType = 1
            
            cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
            cmdSave.ShowSave
            If cmdSave.FileName = "" Then
               Exit Sub
            End If
            Reporte.ExportOptions.DiskFileName = cmdSave.FileName
            Reporte.Export False
   End Select
End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.Value > dtpFecFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub Form_Load()
    Centrar Me
    LlenaCombos
    dtpFecIni.Value = Date
    dtpFecFin.Value = Date + 1
    cmdBusca(0).Enabled = False
    cmdBusca(1).Enabled = False
    chkCliente_Click
    chkEstado_Click
    chkTransportista_Click
    chkMotivo_Click
End Sub

Sub LlenaCombos()
    With cboMotivo
        Isql = "SELECT Codigo, Descripcion FROM vmotivotraslado where lActivo=1"
        Set RsMotivo = Lib.OpenRecordset(Isql, Cn)
        Set .RowSource = RsMotivo
        .DataField = "Descripcion"
        .ListField = "Descripcion"
        .BoundColumn = "Codigo"
    End With

    With cboEstado
        Isql = "SELECT Codigo, Descripcion FROM vEstadoDocumento where lActivo=1"
        Set RsEstado = Lib.OpenRecordset(Isql, Cn)
        Set .RowSource = RsEstado
        .DataField = "Descripcion"
        .ListField = "Descripcion"
        .BoundColumn = "Codigo"
    End With
End Sub

Public Sub Genera()
    Screen.MousePointer = vbHourglass
    Isql = "SELECT tGuia, tDocumento, Destinatario AS Cliente, Transportista, Estado, Motivo, fFecha " & _
           "FROM vGuia WHERE " & sCriterio
    Set RsReporte = Lib.OpenRecordset(Isql, Cn)
    
    Reporte.DiscardSavedData
    Reporte.Database.SetDataSource RsReporte
    Reporte.ReportTitle = "Del " & dtpFecIni.Value & " Al " & dtpFecFin.Value
    Reporte.Text28.SetText sRazonSocial
    frmEmite.CRViewer.DisplayGroupTree = False
    frmEmite.CRViewer.ReportSource = Reporte

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsReporte = Nothing
    Set RsMotivo = Nothing
    Set RsEstado = Nothing
End Sub

