VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepCancelacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cancelaciones"
   ClientHeight    =   5070
   ClientLeft      =   2895
   ClientTop       =   2190
   ClientWidth     =   9045
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepCancelacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   9045
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
      Index           =   2
      Left            =   4522
      Picture         =   "frmRepCancelacion.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4410
      Visible         =   0   'False
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
      Height          =   4320
      Left            =   45
      TabIndex        =   17
      Top             =   15
      Width           =   8970
      Begin VB.CheckBox chkincldoc 
         Caption         =   "Filtrar por Fecha de Pago"
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
         Left            =   6480
         TabIndex        =   25
         Top             =   360
         Width           =   2235
      End
      Begin VB.CheckBox chkDiaContable 
         Caption         =   "Evaluar por Dia Contable"
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
         Left            =   2520
         TabIndex        =   12
         Top             =   3840
         Width           =   3675
      End
      Begin VB.Frame Frame2 
         Caption         =   " Tipo de Reporte "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Left            =   2565
         TabIndex        =   24
         Top             =   2520
         Width           =   3615
         Begin VB.OptionButton optOpcion 
            Caption         =   "Resumido"
            Height          =   240
            Index           =   1
            Left            =   210
            TabIndex        =   11
            Top             =   780
            Width           =   2835
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Detallado"
            Height          =   240
            Index           =   0
            Left            =   210
            TabIndex        =   10
            Top             =   390
            Value           =   -1  'True
            Width           =   3015
         End
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
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1110
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
         Left            =   5460
         Picture         =   "frmRepCancelacion.frx":082E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1080
         Width           =   765
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
         Left            =   6465
         TabIndex        =   6
         Top             =   1125
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin VB.CheckBox chkTipoDocumento 
         Caption         =   "Todos los Documentos"
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
         Left            =   6465
         TabIndex        =   8
         Top             =   1605
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   8325
         Top             =   270
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "Xls"
      End
      Begin VB.ComboBox cboOrden 
         Height          =   360
         ItemData        =   "frmRepCancelacion.frx":0930
         Left            =   2580
         List            =   "frmRepCancelacion.frx":093D
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2025
         Width           =   1860
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   330
         Left            =   2580
         TabIndex        =   2
         Top             =   645
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         _Version        =   393216
         Format          =   67043329
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   330
         Left            =   2580
         TabIndex        =   0
         Top             =   225
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         _Version        =   393216
         Format          =   67043329
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   330
         Left            =   4620
         TabIndex        =   1
         Top             =   225
         Width           =   1590
         _ExtentX        =   2805
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
         Format          =   67043331
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   330
         Left            =   4620
         TabIndex        =   3
         Top             =   645
         Width           =   1590
         _ExtentX        =   2805
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
         Format          =   67043331
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboTipoDocumento 
         Height          =   315
         Left            =   2580
         TabIndex        =   7
         Top             =   1560
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
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
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
         Left            =   1725
         TabIndex        =   23
         Top             =   1170
         Width           =   720
      End
      Begin VB.Label Label 
         Caption         =   "( Factura = Nota de Crédito )"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   6855
         TabIndex        =   22
         Top             =   1845
         Width           =   1815
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Documento :"
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
         Left            =   915
         TabIndex        =   21
         Top             =   1620
         Width           =   1530
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
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
         Left            =   1320
         TabIndex        =   20
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
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
         Index           =   0
         Left            =   1215
         TabIndex        =   19
         Top             =   300
         Width           =   1230
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenado por :"
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
         Index           =   2
         Left            =   1155
         TabIndex        =   18
         Top             =   2115
         Width           =   1290
      End
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
      Index           =   3
      Left            =   5977
      Picture         =   "frmRepCancelacion.frx":095E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4410
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
      Left            =   3067
      Picture         =   "frmRepCancelacion.frx":0A50
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4410
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
      Left            =   1612
      Picture         =   "frmRepCancelacion.frx":0F82
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4410
      Width           =   1455
   End
End
Attribute VB_Name = "frmRepCancelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTipoDocumento As Recordset
Dim RsSumas As Recordset

Dim sCriterio As String
Dim xCriterio As String
Dim sReporte As String

Dim RsCliente As Recordset
Dim RsCompania As Recordset

Dim Resumido As New dsrCancelacionR
Dim Detallado As New dsrCancelacionD

Dim tTarjeta(24) As String
Dim nTarjeta(24) As Double
Dim nPropina(24) As Double
Dim nFactor(24) As Double

Dim tOtro(20) As String
Dim xOtro(20) As String
Dim nOtro(20) As Double

Dim nTarjetaT As Double
Dim nPropinaT As Double
Dim nOtroT As Double
Dim i As Integer

Dim nEfectivoN As Double
Dim nEfectivoE As Double
Dim nEfectivoT As Double
Dim nDolar As Double
Dim nDolar2 As Double
Dim nDolar3 As Double
Dim nCheque As Double
Dim nChequeMN As Double
Dim nChequeME As Double
Dim nTTarjeta As Double
Dim nTOtro As Double
Dim nPunto As Double
Dim nCortesia As Double
Dim nEfeN As Double
Dim nEfeE As Double
Dim nCobrar As Double

Dim fInicio As Date
Dim fFinal As Date

'CESAR DIA CONTABLE
Dim sTitulo As String


Sub LlenaCombos()
   With cboTipoDocumento
      'Compania
      Isql = "Select * from vTipoDocumento where Codigo<>'00' Order By Codigo"
      Set RsTipoDocumento = Lib.OpenRecordset(Isql, Cn)
      Set .RowSource = RsTipoDocumento
      .DataField = "Descripcion"
      .ListField = "Descripcion"
      .BoundColumn = "Codigo"
   End With
End Sub

'CESAR DIA CONTABLE
Private Sub chkDiaContable_Click()
    If chkDiaContable.Value = 1 Then
        Me.dtpHorFin.Visible = False
        Me.dtpHorIni.Visible = False
    Else
        Me.dtpHorFin.Visible = True
        Me.dtpHorIni.Visible = True
    End If
End Sub
'FIN


Private Sub cmdBusca_Click()
   Dim xCriterio As String
   Isql = "Select tCodigoCliente as Codigo, tIdentidad as Identidad, tEmpresa as Descripcion from TCLIENTE order by tEmpresa"
   frmBusca.cboCriterio.Enabled = True
   frmBusca.nPredeterm = 2
   Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                          "Identificador", 2, "Identidad", 1500, 0, 0, "", _
                                          "Razón Comercial", 2, "Descripcion", 4500, 0, 0, "")
   frmBusca.Show vbModal
   If Not wEnter Then
      Exit Sub
   End If
   sCliente = sCodigo
   txtCliente.Text = sDescrip
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
      
   If Index = 3 Then
      Unload Me
      Exit Sub
   End If
   
   sCriterio = ""
   xCriterio = ""

'CESAR DIA CONTABLE
If chkDiaContable.Value = 0 Then
   If dtpFecIni.Value + dtpHorIni.Value > dtpFecFin.Value + dtpHorFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      Exit Sub
      dtpFecFin.SetFocus
   End If
Else
   If dtpFecIni.Value > dtpFecFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      Exit Sub
      dtpFecFin.SetFocus
   End If
End If
'FIN CESAR DIA CONTABLE

   If chkCliente.Value = 0 Then
      If sCliente = "" Then
         MsgBox "Debe escoger el Cliente", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " AND tCodigoCliente ='" & sCliente & "'"
      xCriterio = xCriterio & " AND tCodigoCliente ='" & sCliente & "'"
   End If
   
   If chkTipoDocumento.Value = 0 Then
      If cboTipoDocumento.Text = "" Then
         MsgBox "Debe escoger un Tipo de Documento", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " AND tTipoDocumento ='" & cboTipoDocumento.BoundText & "'"
      xCriterio = xCriterio & " AND Codigo ='" & cboTipoDocumento.BoundText & "'"
   End If
            
   cmdOpcion(Index).Enabled = False
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If RsSumas.EOF = True Then
                  Screen.MousePointer = vbDefault
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  cmdOpcion(Index).Enabled = True
                  Exit Sub
               End If
               
               If optOpcion(0).Value = True Then
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado.PaperOrientation = crPortrait
               Else
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crPortrait
               End If
               
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
          
          Case Is = 1 ' Imprimir
               Genera
               Screen.MousePointer = vbDefault
               If RsSumas.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  cmdOpcion(Index).Enabled = True
                  Exit Sub
               End If
               If optOpcion(0).Value = True Then
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado.PaperOrientation = crLandscape
                  Detallado.PrintOut
               Else
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crPortrait
                  Resumido.PrintOut
               End If
          
          Case Is = 2 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If RsSumas.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  cmdOpcion(Index).Enabled = True
                  Exit Sub
               End If
               Screen.MousePointer = vbDefault
               'On Error GoTo Cancelar

               If optOpcion(0).Value = True Then
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado.PaperOrientation = crLandscape
                  
                  Detallado.ExportOptions.FormatType = 21
                  Detallado.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     cmdOpcion(Index).Enabled = True
                     Exit Sub
                  End If
                  Detallado.ExportOptions.DiskFileName = cmdSave.FileName
                  Detallado.Export False
               Else
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crPortrait
                  Resumido.ExportOptions.FormatType = 21
                  Resumido.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     cmdOpcion(Index).Enabled = True
                     Exit Sub
                  End If
                  Resumido.ExportOptions.DiskFileName = cmdSave.FileName
                  Resumido.Export False
               End If
         
          Case Is = 3 ' Salir
               Unload Me
   End Select
   cmdOpcion(Index).Enabled = True
Cancelar:
   Exit Sub
End Sub

Private Sub chkCliente_Click()
   If chkCliente.Value = 1 Then
      sCliente = ""
      txtCliente.Text = ""
      cmdBusca.Enabled = False
   Else
      cmdBusca.Enabled = True
   End If
End Sub

Private Sub chkTipoDocumento_Click()
   If chkTipoDocumento.Value = 1 Then
      cboTipoDocumento.Enabled = False
      cboTipoDocumento.Text = ""
   Else
      cboTipoDocumento.Enabled = True
   End If
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
   dtpFecFin.Value = Date
   cmdBusca.Enabled = False
   cboOrden.Text = cboOrden.List(0)
    
   cboTipoDocumento.Enabled = False

   sReporte = dbTemporal(sCaja, 14, "tDocumento", "varchar(15)", "fRegistro", "datetime", "nEfectivoSoles", "float", "nVenta", "float", _
                        "nEfectivoMN", "float", "nEfectivoME", "float", "nTarjeta", "float", "nPropina", "float", "nCheque", "float", "nOtro", "float", _
                        "nPunto", "float", "nCortesia", "float", "nCobrar", "float", "tEstadoDocumento", "Varchar(2)")
                        
End Sub
Public Sub Genera()
   Dim oComando As clsComando
   Dim sOrden As String
   Screen.MousePointer = vbHourglass
      
   Cn.Execute "delete from " & sReporte
   Select Case cboOrden.Text
          Case Is = "Correlativo"
               sOrden = "tDocumento"
          Case Is = "Montos"
               sOrden = "nVenta"
          Case Is = "Fechas"
               sOrden = "year(" & sReporte & ".fregistro), month(" & sReporte & ".fregistro), day(" & sReporte & ".fregistro), tdocumento"
          Case Else
               sOrden = "tDocumento"
   End Select
     Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_Cancelacion", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    
    'CESAR DIA CONTABLE
    If chkDiaContable.Value = 0 Then
        fInicio = Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm")
        fFinal = Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm")
    Else
        fInicio = Format(dtpFecIni.Value, "yyyy/mm/dd")
        fFinal = Format(dtpFecFin.Value, "yyyy/mm/dd")
    End If
    'FIN CESAR DIA CONTABLE
    
    oComando.CreateParameter "@flagTipo", adBoolean, adParamInput, 1, optOpcion(0).Value
    oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 20, sCliente
    oComando.CreateParameter "@tTipoDoc", adVarChar, adParamInput, 20, Me.cboTipoDocumento.BoundText
    oComando.CreateParameter "@sOrden", adVarChar, adParamInput, 140, Me.cboOrden.Text
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@diaContable", adBoolean, adParamInput, 1, Me.chkDiaContable.Value
    oComando.CreateParameter "@fpago", adBoolean, adParamInput, 1, Me.chkincldoc.Value
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
  
   If chkDiaContable.Value = 0 Then
        nDolar = Calcular("select sum(nDolar) as Codigo " & _
                          "FROM dbo.MDOCUMENTO INNER JOIN dbo.DPAGODOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DPAGODOCUMENTO.tDocumento " & _
                          "where tTipoPago='01' and tMoneda='02' and tTipoDocumento <> '00' and tEstadoDocumento='02' and MDOCUMENTO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:nn") & "' and MDOCUMENTO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:nn") & "' " & sCriterio, Cn)
                          
        nDolar2 = Calcular("select sum(nDolar) as Codigo " & _
                          "FROM dbo.MDOCUMENTO INNER JOIN dbo.DPAGODOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DPAGODOCUMENTO.tDocumento " & _
                          "where tTipoPago='01' and tMoneda='03' and tTipoDocumento <> '00' and tEstadoDocumento='02' and MDOCUMENTO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:nn") & "' and MDOCUMENTO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:nn") & "' " & sCriterio, Cn)
                          
        nDolar3 = Calcular("select sum(nDolar) as Codigo " & _
                          "FROM dbo.MDOCUMENTO INNER JOIN dbo.DPAGODOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DPAGODOCUMENTO.tDocumento " & _
                          "where tTipoPago='01' and tMoneda='04' and tTipoDocumento <> '00' and tEstadoDocumento='02' and MDOCUMENTO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:nn") & "' and MDOCUMENTO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:nn") & "' " & sCriterio, Cn)
                 
        nEfectivoN = Calcular("select sum(nMonto) as Codigo " & _
                              "FROM dbo.MDOCUMENTO INNER JOIN dbo.DPAGODOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DPAGODOCUMENTO.tDocumento " & _
                              "where tTipoPago='01' and tMoneda='01' and tTipoDocumento <> '00' and tEstadoDocumento='02' and MDOCUMENTO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:nn") & "' and MDOCUMENTO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:nn") & "' " & sCriterio, Cn)
                 
        nEfectivoE = Calcular("select sum((nDolar-nMonto)*nTipoCambio) as Codigo " & _
                              "FROM dbo.MDOCUMENTO INNER JOIN dbo.DPAGODOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DPAGODOCUMENTO.tDocumento " & _
                              "where tTipoPago='01' and tMoneda in ('02','03','04') and tTipoDocumento <> '00' and tEstadoDocumento='02' and MDOCUMENTO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:nn") & "' and MDOCUMENTO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:nn") & "' " & sCriterio, Cn)
   Else
        nDolar = Calcular("select sum(nDolar) as Codigo " & _
                          "FROM dbo.MDOCUMENTO INNER JOIN dbo.DPAGODOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DPAGODOCUMENTO.tDocumento " & _
                          "where tTipoPago='01' and tMoneda='02' and tTipoDocumento <> '00' and tEstadoDocumento='02' and MDOCUMENTO.fDiaContable >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & "' and MDOCUMENTO.fDiaContable <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & "' " & sCriterio, Cn)
                          
        nDolar2 = Calcular("select sum(nDolar) as Codigo " & _
                          "FROM dbo.MDOCUMENTO INNER JOIN dbo.DPAGODOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DPAGODOCUMENTO.tDocumento " & _
                          "where tTipoPago='01' and tMoneda='03' and tTipoDocumento <> '00' and tEstadoDocumento='02' and MDOCUMENTO.fDiaContable >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & "' and MDOCUMENTO.fDiaContable <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & "' " & sCriterio, Cn)
                          
        nDolar3 = Calcular("select sum(nDolar) as Codigo " & _
                          "FROM dbo.MDOCUMENTO INNER JOIN dbo.DPAGODOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DPAGODOCUMENTO.tDocumento " & _
                          "where tTipoPago='01' and tMoneda='04' and tTipoDocumento <> '00' and tEstadoDocumento='02' and MDOCUMENTO.fDiaContable >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & "' and MDOCUMENTO.fDiaContable <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & "' " & sCriterio, Cn)
                 
        nEfectivoN = Calcular("select sum(nMonto) as Codigo " & _
                              "FROM dbo.MDOCUMENTO INNER JOIN dbo.DPAGODOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DPAGODOCUMENTO.tDocumento " & _
                              "where tTipoPago='01' and tMoneda='01' and tTipoDocumento <> '00' and tEstadoDocumento='02' and MDOCUMENTO.fDiaContable >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & "' and MDOCUMENTO.fDiaContable <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & "' " & sCriterio, Cn)
                 
        nEfectivoE = Calcular("select sum((nDolar-nMonto)*nTipoCambio) as Codigo " & _
                              "FROM dbo.MDOCUMENTO INNER JOIN dbo.DPAGODOCUMENTO ON dbo.MDOCUMENTO.tDocumento = dbo.DPAGODOCUMENTO.tDocumento " & _
                              "where tTipoPago='01' and tMoneda in ('02','03','04') and tTipoDocumento <> '00' and tEstadoDocumento='02' and MDOCUMENTO.fDiaContable >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & "' and MDOCUMENTO.fDiaContable <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & "' " & sCriterio, Cn)
   End If
   
   If chkDiaContable.Value = 0 Then
         sTitulo = "Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.Value, "HH:mm") & " Hrs Al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.Value, "HH:mm") & " Hrs"
   Else
         sTitulo = "Por Dia Contable Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.Value, "dd/mmm/yyyy")
   End If
                        
   If optOpcion(0).Value = True Then
      
      Set RsSumas = oComando.GetSP()
      Detallado.DiscardSavedData
      Detallado.Database.SetDataSource RsSumas

      Detallado.ReportTitle = sTitulo
      Detallado.Text5.SetText sRazonSocial
      Detallado.EfectivoN.SetText Format(nEfectivoN - nEfectivoE, "###,##0.00")
      Detallado.EfectivoE.SetText Format(nDolar, "###,##0.00")
      Detallado.txtNameMN.SetText sMonedaN
      Detallado.txtNameME.SetText sMonedaE
      If pais = "003" Then
        Detallado.txtNameME2.SetText sMonedaE2
        Detallado.txtNameME3.SetText sMonedaE3
        Detallado.EfectivoE2.SetText Format(nDolar2, "###,##0.00")
        Detallado.EfectivoE3.SetText Format(nDolar3, "###,##0.00")
      Else
        Detallado.txtNameME2.SetText ""
        Detallado.txtNameME3.SetText ""
        Detallado.EfectivoE2.SetText ""
        Detallado.EfectivoE3.SetText ""
      End If
      'Detallado.Text30.SetText sMonN
      'Detallado.Text31.SetText sMonE
      Detallado.Text32.SetText localConectado
      frmEmite.CRViewer.DisplayGroupTree = False
      frmEmite.CRViewer.ReportSource = Detallado
      
   Else
      'Tarjetas
      Set RsSumas = Lib.OpenRecordset("Select tCodigoTarjeta, tDetallado, nFactorRetencion as nFactor from TTARJETACREDITO where lActivo=1", Cn)
      For i = 1 To 24
          RsSumas.Filter = "tCodigoTarjeta='" & Right("00" & LTrim(str(i)), 2) & "'"
          If RsSumas.RecordCount > 0 Then
             tTarjeta(i) = IIf(IsNull(RsSumas!tDetallado), 0, RsSumas!tDetallado)
             nFactor(i) = IIf(IsNull(RsSumas!nFactor), 0, RsSumas!nFactor)
          Else
             tTarjeta(i) = "(no definido)"
             nFactor(i) = 0
          End If
      Next i
      
      'Sumas Tarjetas
      If chkDiaContable.Value = 0 Then
            Isql = "Select tTarjeta, sum(nMonto) as nMonto, sum(DPAGODOCUMENTO.nPropina) as nPropina " & _
                   "FROM dbo.DPAGODOCUMENTO INNER JOIN dbo.MDOCUMENTO ON dbo.DPAGODOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
                   "Where tTipoPago='02' and tTipoDocumento <> '00' and tEstadoDocumento='02' and MDOCUMENTO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:nn") & "' and MDOCUMENTO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:nn") & "' " & sCriterio & " " & _
                   "Group by tTarjeta order by tTarjeta"
      Else
            Isql = "Select tTarjeta, sum(nMonto) as nMonto, sum(DPAGODOCUMENTO.nPropina) as nPropina " & _
                   "FROM dbo.DPAGODOCUMENTO INNER JOIN dbo.MDOCUMENTO ON dbo.DPAGODOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
                   "Where tTipoPago='02' and tTipoDocumento <> '00' and tEstadoDocumento='02' and MDOCUMENTO.fDiaContable >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & "' and MDOCUMENTO.fDiaContable <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & "' " & sCriterio & " " & _
                   "Group by tTarjeta order by tTarjeta"
      End If
             
      Set RsSumas = Lib.OpenRecordset(Isql, Cn)
      
      For i = 1 To 24
          RsSumas.Filter = "tTarjeta='" & Right("00" & LTrim(str(i)), 2) & "'"
          If RsSumas.RecordCount > 0 Then
             nTarjeta(i) = IIf(IsNull(RsSumas!nMonto), 0, RsSumas!nMonto)
             nPropina(i) = IIf(IsNull(RsSumas!nPropina), 0, RsSumas!nPropina)
          Else
             nTarjeta(i) = 0
             nPropina(i) = 0
          End If
          nTarjetaT = nTarjetaT + nTarjeta(i)
          nPropinaT = nPropinaT + nPropina(i)
      Next i
      
      'Otros Tipos de Pago
      Isql = "select * from vTipoCancelacion where lActivo=1 Order by Codigo"
      Set RsSumas = Lib.OpenRecordset(Isql, Cn)
      If RsSumas.RecordCount > 0 Then
         RsSumas.MoveFirst
         For i = 1 To 20
             If RsSumas.EOF Then
                tOtro(i) = "( no definido )"
                xOtro(i) = ""
             Else
                tOtro(i) = RsSumas!Descripcion
                xOtro(i) = RsSumas!codigo
                RsSumas.MoveNext
             End If
         Next i
      End If

      'Sumas Otros Tipos
      If chkDiaContable.Value = 0 Then
        Isql = "Select tOtroTipoPago, sum(nMonto) as nMonto " & _
               "FROM dbo.DPAGODOCUMENTO INNER JOIN dbo.MDOCUMENTO ON dbo.DPAGODOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
               "Where tTipoPago='04' and tTipoDocumento <> '00' and tEstadoDocumento='02' and MDOCUMENTO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:nn") & "' and MDOCUMENTO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:nn") & "' " & sCriterio & " " & _
               "Group by tOtroTipoPago order by tOtroTipoPago"
      Else
        Isql = "Select tOtroTipoPago, sum(nMonto) as nMonto " & _
               "FROM dbo.DPAGODOCUMENTO INNER JOIN dbo.MDOCUMENTO ON dbo.DPAGODOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento " & _
               "Where tTipoPago='04' and tTipoDocumento <> '00' and tEstadoDocumento='02' and MDOCUMENTO.fDiaContable >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & "' and MDOCUMENTO.fDiaContable <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & "' " & sCriterio & " " & _
               "Group by tOtroTipoPago order by tOtroTipoPago"
      End If
             
      Set RsSumas = Lib.OpenRecordset(Isql, Cn)
      
      For i = 1 To 20
          RsSumas.Filter = "tOtroTipoPago='" & xOtro(i) & "'"
          If RsSumas.RecordCount > 0 Then
             nOtro(i) = IIf(IsNull(RsSumas!nMonto), 0, RsSumas!nMonto)
          Else
             nOtro(i) = 0
          End If
          nOtroT = nOtroT + nOtro(i)
      Next i
      
      Resumido.Text73.SetText sMonN
      Resumido.txtME.SetText sMonE
      If pais = "003" Then
        Resumido.txtME2.SetText sMonE2
        Resumido.txtME3.SetText sMonE3
      Else
        Resumido.txtME2.SetText ""
        Resumido.txtME3.SetText ""
      End If
      Resumido.Text19.SetText sMonN
      Resumido.Text20.SetText sMonE
      Resumido.Text17.SetText localConectado
      If chkDiaContable.Value = 0 Then
        nCobrar = Calcular("select sum(MDOCUMENTO.nVenta) as Codigo FROM MDOCUMENTO where MDOCUMENTO.tEstadoDocumento='03' and  MDOCUMENTO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:nn") & "' and MDOCUMENTO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:nn") & "' " & sCriterio, Cn)
      Else
        nCobrar = Calcular("select sum(MDOCUMENTO.nVenta) as Codigo FROM MDOCUMENTO where MDOCUMENTO.tEstadoDocumento='03' and  MDOCUMENTO.fDiaContable >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & "' and MDOCUMENTO.fDiaContable <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & "' " & sCriterio, Cn)
      End If
      
      Set RsSumas = oComando.GetSP()
      nEfeN = IIf(IsNull(RsSumas!nEfectivoMN), 0, RsSumas!nEfectivoMN)
      nEfeE = IIf(IsNull(RsSumas!nEfectivoME), 0, RsSumas!nEfectivoME)
      nEfectivoT = IIf(IsNull(RsSumas!nEfectivoSoles), 0, RsSumas!nEfectivoSoles) + IIf(IsNull(RsSumas!nEfectivoMN), 0, RsSumas!nEfectivoMN) - IIf(IsNull(RsSumas!nvuelto), 0, RsSumas!nvuelto)
      nEfectivoN = IIf(IsNull(RsSumas!nEfectivoMN), 0, RsSumas!nEfectivoMN) - nEfectivoE
      nEfectivoE = nDolar
      
      nCheque = IIf(IsNull(RsSumas!nChequeMN), 0, RsSumas!nChequeMN) + IIf(IsNull(RsSumas!nChequeSoles), 0, RsSumas!nChequeSoles)
      nChequeMN = IIf(IsNull(RsSumas!nChequeMN), 0, RsSumas!nChequeMN)
      nChequeME = IIf(IsNull(RsSumas!nChequeME), 0, RsSumas!nChequeME)
      nPunto = IIf(IsNull(RsSumas!nPunto), 0, RsSumas!nPunto)
      
      If chkDiaContable.Value = 0 Then
        nCortesia = Calcular("select sum(nVenta) as Codigo from MDOCUMENTO where tEstadoDocumento='02' and  ttipodocumento <>'00' and tCortesia <>'' and MDOCUMENTO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:nn") & "' and MDOCUMENTO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:nn") & "' " & sCriterio, Cn)
      Else
        nCortesia = Calcular("select sum(nVenta) as Codigo from MDOCUMENTO where tEstadoDocumento='02' and ttipodocumento <>'00' and tCortesia <>'' and MDOCUMENTO.fDiaContable >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & "' and MDOCUMENTO.fDiaContable <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & "' " & sCriterio, Cn)
      End If
      
      'nTTarjeta = nTarjeta(1) + nTarjeta(2) + nTarjeta(3) + nTarjeta(4) + nTarjeta(5) + nTarjeta(6) + nTarjeta(7) + nTarjeta(8)
      
      nTOtro = nOtro(1) + nOtro(2) + nOtro(3) + nOtro(4) + nOtro(5) + nOtro(6) + nOtro(7) + nOtro(8) + nOtro(9) + nOtro(10) + nOtro(11) + nOtro(12) + nOtro(13) + nOtro(14) + nOtro(15) + nOtro(16) + nOtro(17) + nOtro(18) + nOtro(19) + nOtro(20)
                              
      Resumido.TotMN.SetText Format(nEfectivoN, "###,###,#0.00")
      Resumido.TotME.SetText Format(nEfectivoE, "###,###,#0.00")
      If pais = "003" Then
        Resumido.TotME2.SetText Format(nDolar2, "###,###,#0.00")
        Resumido.TotME3.SetText Format(nDolar3, "###,###,#0.00")
        Resumido.EFEME2.SetText Format(nDolar2, "###,###,#0.00")
        Resumido.EFEME3.SetText Format(nDolar3, "###,###,#0.00")
      Else
        Resumido.TotME2.SetText ""
        Resumido.TotME3.SetText ""
        Resumido.EFEME2.SetText ""
        Resumido.EFEME3.SetText ""
      End If
      Resumido.EFEMN.SetText Format(nEfeN, "###,###,#0.00")
      Resumido.EFEME.SetText Format(nEfeE, "###,###,#0.00")
      Resumido.Cobrar.SetText Format(nCobrar, "###,###,#0.00")
      Resumido.TotEfectivo.SetText Format(nEfectivoT, "###,###,##0.00")
      Resumido.TotCheque.SetText Format(nCheque, "###,###,#0.00")
      Resumido.TotChequeMN.SetText Format(nChequeMN, "###,###,#0.00")
      Resumido.TotChequeME.SetText Format(nChequeME, "###,###,#0.00")
      Resumido.TotPunto.SetText Format(nPunto, "###,###,#0.00")
      Resumido.TotCortesia.SetText Format(nCortesia, "###,###,##0.00")
      Resumido.Venta.SetText Format(IIf(IsNull(RsSumas!Venta), 0, RsSumas!Venta), "###,###,##0.00")
      
      Dim nNC As Double
      If chkDiaContable.Value = 0 Then
        nNC = Calcular("select sum(nVenta) as Codigo from MNOTACREDITO where MNOTACREDITO.tEstadoDocumento='02' and  MNOTACREDITO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:nn") & "' and MNOTACREDITO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:nn") & "'", Cn)
      Else
        nNC = Calcular("select sum(nVenta) as Codigo from MNOTACREDITO where MNOTACREDITO.tEstadoDocumento='02' and  MNOTACREDITO.fDiaContable >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & "' and MNOTACREDITO.fDiaContable <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & "'", Cn)
      End If
      
      Resumido.txtNC.SetText Format(nNC, "###,###,##0.00")
      Resumido.txtRV.SetText Format(IIf(IsNull(RsSumas!Venta), 0, RsSumas!Venta) - nNC - nCortesia, "###,###,##0.00")
         
      'Tarjetas
      Resumido.Tarjeta1.SetText tTarjeta(1)
      Resumido.Tarjeta2.SetText tTarjeta(2)
      Resumido.Tarjeta3.SetText tTarjeta(3)
      Resumido.Tarjeta4.SetText tTarjeta(4)
      Resumido.Tarjeta5.SetText tTarjeta(5)
      Resumido.Tarjeta6.SetText tTarjeta(6)
      Resumido.Tarjeta7.SetText tTarjeta(7)
      Resumido.Tarjeta8.SetText tTarjeta(8)
      Resumido.Tarjeta9.SetText tTarjeta(9)
      Resumido.Tarjeta10.SetText tTarjeta(10)
      Resumido.Tarjeta11.SetText tTarjeta(11)
      Resumido.Tarjeta12.SetText tTarjeta(12)
      Resumido.Tarjeta13.SetText tTarjeta(13)
      Resumido.Tarjeta14.SetText tTarjeta(14)
      Resumido.Tarjeta15.SetText tTarjeta(15)
      Resumido.Tarjeta16.SetText tTarjeta(16)
      Resumido.Tarjeta17.SetText tTarjeta(17)
      Resumido.Tarjeta18.SetText tTarjeta(18)
      Resumido.Tarjeta19.SetText tTarjeta(19)
      Resumido.Tarjeta20.SetText tTarjeta(20)
      Resumido.Tarjeta21.SetText tTarjeta(21)
      Resumido.Tarjeta22.SetText tTarjeta(22)
      Resumido.Tarjeta23.SetText tTarjeta(23)
      Resumido.Tarjeta24.SetText tTarjeta(24)
           
      Resumido.TotTC1.SetText Format(nTarjeta(1), "###,###,#0.00")
      Resumido.TotTC2.SetText Format(nTarjeta(2), "###,###,#0.00")
      Resumido.TotTC3.SetText Format(nTarjeta(3), "###,###,#0.00")
      Resumido.TotTC4.SetText Format(nTarjeta(4), "###,###,#0.00")
      Resumido.TotTC5.SetText Format(nTarjeta(5), "###,###,#0.00")
      Resumido.TotTC6.SetText Format(nTarjeta(6), "###,###,#0.00")
      Resumido.TotTC7.SetText Format(nTarjeta(7), "###,###,#0.00")
      Resumido.TotTC8.SetText Format(nTarjeta(8), "###,###,#0.00")
      Resumido.TotTC9.SetText Format(nTarjeta(9), "###,###,#0.00")
      Resumido.TotTC10.SetText Format(nTarjeta(10), "###,###,#0.00")
      Resumido.TotTC11.SetText Format(nTarjeta(11), "###,###,#0.00")
      Resumido.TotTC12.SetText Format(nTarjeta(12), "###,###,#0.00")
      Resumido.TotTC13.SetText Format(nTarjeta(13), "###,###,#0.00")
      Resumido.TotTC14.SetText Format(nTarjeta(14), "###,###,#0.00")
      Resumido.TotTC15.SetText Format(nTarjeta(15), "###,###,#0.00")
      Resumido.TotTC16.SetText Format(nTarjeta(16), "###,###,#0.00")
      Resumido.TotTC17.SetText Format(nTarjeta(17), "###,###,#0.00")
      Resumido.TotTC18.SetText Format(nTarjeta(18), "###,###,#0.00")
      Resumido.TotTC19.SetText Format(nTarjeta(19), "###,###,#0.00")
      Resumido.TotTC20.SetText Format(nTarjeta(20), "###,###,#0.00")
      Resumido.TotTC21.SetText Format(nTarjeta(21), "###,###,#0.00")
      Resumido.TotTC22.SetText Format(nTarjeta(22), "###,###,#0.00")
      Resumido.TotTC23.SetText Format(nTarjeta(23), "###,###,#0.00")
      Resumido.TotTC24.SetText Format(nTarjeta(24), "###,###,#0.00")
      
      nTTarjeta = 0
      For i = 1 To 24
        nTTarjeta = nTTarjeta + nTarjeta(i)
      Next i
      Resumido.TotTC.SetText Format(nTTarjeta, "###,###,##0.00")
      
      Resumido.Prop1.SetText Format(nPropina(1), "###,###,#0.00")
      Resumido.Prop2.SetText Format(nPropina(2), "###,###,#0.00")
      Resumido.Prop3.SetText Format(nPropina(3), "###,###,#0.00")
      Resumido.Prop4.SetText Format(nPropina(4), "###,###,#0.00")
      Resumido.Prop5.SetText Format(nPropina(5), "###,###,#0.00")
      Resumido.Prop6.SetText Format(nPropina(6), "###,###,#0.00")
      Resumido.Prop7.SetText Format(nPropina(7), "###,###,#0.00")
      Resumido.Prop8.SetText Format(nPropina(8), "###,###,#0.00")
      Resumido.Prop9.SetText Format(nPropina(9), "###,###,#0.00")
      Resumido.Prop10.SetText Format(nPropina(10), "###,###,#0.00")
      Resumido.Prop11.SetText Format(nPropina(11), "###,###,#0.00")
      Resumido.Prop12.SetText Format(nPropina(12), "###,###,#0.00")
      Resumido.Prop13.SetText Format(nPropina(13), "###,###,#0.00")
      Resumido.Prop14.SetText Format(nPropina(14), "###,###,#0.00")
      Resumido.Prop15.SetText Format(nPropina(15), "###,###,#0.00")
      Resumido.Prop16.SetText Format(nPropina(16), "###,###,#0.00")
      Resumido.Prop17.SetText Format(nPropina(17), "###,###,#0.00")
      Resumido.Prop18.SetText Format(nPropina(18), "###,###,#0.00")
      Resumido.Prop19.SetText Format(nPropina(19), "###,###,#0.00")
      Resumido.Prop20.SetText Format(nPropina(20), "###,###,#0.00")
      Resumido.Prop21.SetText Format(nPropina(21), "###,###,#0.00")
      Resumido.Prop22.SetText Format(nPropina(22), "###,###,#0.00")
      Resumido.Prop23.SetText Format(nPropina(23), "###,###,#0.00")
      Resumido.Prop24.SetText Format(nPropina(24), "###,###,#0.00")
      
      'Resumido.TotProp.SetText Format(nPropina(1) + nPropina(2) + nPropina(3) + nPropina(4) + nPropina(5) + nPropina(6) + nPropina(7) + nPropina(8), "###,###,##0.00")
    
      Dim totalprod As Double
      totalprod = 0
      For i = 1 To 24
        totalprod = totalprod + nPropina(i)
      Next i
      
      Resumido.TotProp.SetText Format(totalprod, "###,###,#0.00")
    
      Resumido.Dep1.SetText Format(nTarjeta(1) + nPropina(1), "###,###,#0.00")
      Resumido.Dep2.SetText Format(nTarjeta(2) + nPropina(2), "###,###,#0.00")
      Resumido.Dep3.SetText Format(nTarjeta(3) + nPropina(3), "###,###,#0.00")
      Resumido.Dep4.SetText Format(nTarjeta(4) + nPropina(4), "###,###,#0.00")
      Resumido.Dep5.SetText Format(nTarjeta(5) + nPropina(5), "###,###,#0.00")
      Resumido.Dep6.SetText Format(nTarjeta(6) + nPropina(6), "###,###,#0.00")
      Resumido.Dep7.SetText Format(nTarjeta(7) + nPropina(7), "###,###,#0.00")
      Resumido.Dep8.SetText Format(nTarjeta(8) + nPropina(8), "###,###,#0.00")
      Resumido.Dep9.SetText Format(nTarjeta(9) + nPropina(9), "###,###,#0.00")
      Resumido.Dep10.SetText Format(nTarjeta(10) + nPropina(10), "###,###,#0.00")
      Resumido.Dep11.SetText Format(nTarjeta(11) + nPropina(11), "###,###,#0.00")
      Resumido.Dep12.SetText Format(nTarjeta(12) + nPropina(12), "###,###,#0.00")
      Resumido.Dep13.SetText Format(nTarjeta(13) + nPropina(13), "###,###,#0.00")
      Resumido.Dep14.SetText Format(nTarjeta(14) + nPropina(14), "###,###,#0.00")
      Resumido.Dep15.SetText Format(nTarjeta(15) + nPropina(15), "###,###,#0.00")
      Resumido.Dep16.SetText Format(nTarjeta(16) + nPropina(16), "###,###,#0.00")
      Resumido.Dep17.SetText Format(nTarjeta(17) + nPropina(17), "###,###,#0.00")
      Resumido.Dep18.SetText Format(nTarjeta(18) + nPropina(18), "###,###,#0.00")
      Resumido.Dep19.SetText Format(nTarjeta(19) + nPropina(19), "###,###,#0.00")
      Resumido.Dep20.SetText Format(nTarjeta(20) + nPropina(20), "###,###,#0.00")
      Resumido.Dep21.SetText Format(nTarjeta(21) + nPropina(21), "###,###,#0.00")
      Resumido.Dep22.SetText Format(nTarjeta(22) + nPropina(22), "###,###,#0.00")
      Resumido.Dep23.SetText Format(nTarjeta(23) + nPropina(23), "###,###,#0.00")
      Resumido.Dep24.SetText Format(nTarjeta(24) + nPropina(24), "###,###,#0.00")
      'Resumido.TotDep.SetText Format(nTarjeta(1) + nTarjeta(2) + nTarjeta(3) + nTarjeta(4) + nTarjeta(5) + nTarjeta(6) + nTarjeta(7) + nTarjeta(8) + nPropina(1) + nPropina(2) + nPropina(3) + nPropina(4) + nPropina(5) + nPropina(6) + nPropina(7) + nPropina(8), "###,###,#0.00")
      
      Dim totalDep As Double
      totalDep = 0
      For i = 1 To 24
        totalDep = totalDep + nTarjeta(i) + nPropina(i)
      Next i
      
     Resumido.TotDep.SetText Format(totalDep, "###,###,#0.00")

      'Otros
      Resumido.Otro1.SetText tOtro(1)
      Resumido.Otro2.SetText tOtro(2)
      Resumido.Otro3.SetText tOtro(3)
      Resumido.Otro4.SetText tOtro(4)
      Resumido.Otro5.SetText tOtro(5)
      Resumido.Otro6.SetText tOtro(6)
      Resumido.Otro7.SetText tOtro(7)
      Resumido.Otro8.SetText tOtro(8)
      Resumido.Otro9.SetText tOtro(9)
      Resumido.Otro10.SetText tOtro(10)
      
      Resumido.Otro11.SetText tOtro(11)
      Resumido.Otro12.SetText tOtro(12)
      Resumido.Otro13.SetText tOtro(13)
      Resumido.Otro14.SetText tOtro(14)
      Resumido.Otro15.SetText tOtro(15)
      Resumido.Otro16.SetText tOtro(16)
      Resumido.Otro17.SetText tOtro(17)
      Resumido.Otro18.SetText tOtro(18)
      Resumido.Otro19.SetText tOtro(19)
      Resumido.Otro20.SetText tOtro(20)
                  
      Resumido.TotOtro1.SetText Format(nOtro(1), "###,###,##0.00")
      Resumido.TotOtro2.SetText Format(nOtro(2), "###,###,##0.00")
      Resumido.TotOtro3.SetText Format(nOtro(3), "###,###,##0.00")
      Resumido.TotOtro4.SetText Format(nOtro(4), "###,###,##0.00")
      Resumido.TotOtro5.SetText Format(nOtro(5), "###,###,##0.00")
      Resumido.TotOtro6.SetText Format(nOtro(6), "###,###,##0.00")
      Resumido.TotOtro7.SetText Format(nOtro(7), "###,###,##0.00")
      Resumido.TotOtro8.SetText Format(nOtro(8), "###,###,##0.00")
      Resumido.TotOtro9.SetText Format(nOtro(9), "###,###,##0.00")
      Resumido.TotOtro10.SetText Format(nOtro(10), "###,###,##0.00")
      Resumido.TotOtro11.SetText Format(nOtro(11), "###,###,##0.00")
      Resumido.TotOtro12.SetText Format(nOtro(12), "###,###,##0.00")
      Resumido.TotOtro13.SetText Format(nOtro(13), "###,###,##0.00")
      Resumido.TotOtro14.SetText Format(nOtro(14), "###,###,##0.00")
      Resumido.TotOtro15.SetText Format(nOtro(15), "###,###,##0.00")
      Resumido.TotOtro16.SetText Format(nOtro(16), "###,###,##0.00")
      Resumido.TotOtro17.SetText Format(nOtro(17), "###,###,##0.00")
      Resumido.TotOtro18.SetText Format(nOtro(18), "###,###,##0.00")
      Resumido.TotOtro19.SetText Format(nOtro(19), "###,###,##0.00")
      Resumido.TotOtro20.SetText Format(nOtro(20), "###,###,##0.00")
      
      
      Resumido.TotOtro.SetText Format(nTOtro, "###,###,##0.00")
            
      Resumido.TotCobrado.SetText Format(nEfectivoT + nCheque + nPunto + nCortesia + nTTarjeta + nTOtro, "###,###,##0.00")
                  
      Resumido.DiscardSavedData
      Resumido.Database.SetDataSource RsSumas
      'CESAR DIA CONTABLE Titulo
      Resumido.ReportTitle = sTitulo
      Resumido.Text5.SetText sRazonSocial
      frmEmite.CRViewer.DisplayGroupTree = False
      frmEmite.CRViewer.ReportSource = Resumido
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsSumas = Nothing
   Set frmRepRegistroVenta = Nothing
   Set RsCliente = Nothing
   Set RsTipoDocumento = Nothing
   
   If sReporte <> "" Then
      Cn.Execute "Drop Table " & sReporte
   End If
End Sub
