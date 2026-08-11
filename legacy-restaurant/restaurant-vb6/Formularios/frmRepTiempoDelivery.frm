VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepTiempoDelivery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tiempos en Delivery"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "frmRepTiempoDelivery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
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
      Left            =   1590
      Picture         =   "frmRepTiempoDelivery.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2460
      UseMaskColor    =   -1  'True
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
      Left            =   3045
      Picture         =   "frmRepTiempoDelivery.frx":053E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2460
      UseMaskColor    =   -1  'True
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
      Left            =   5955
      Picture         =   "frmRepTiempoDelivery.frx":0A70
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2460
      UseMaskColor    =   -1  'True
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
      Height          =   2355
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7455
      Begin VB.OptionButton optCentral 
         Caption         =   "Central Delivery"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4860
         TabIndex        =   14
         Top             =   1920
         Width           =   2235
      End
      Begin VB.OptionButton optToma 
         Caption         =   "Toma Delivery"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1590
         TabIndex        =   13
         Top             =   1920
         Value           =   -1  'True
         Width           =   2010
      End
      Begin VB.CheckBox chkMotorizado 
         Caption         =   "Todos los Motorizados"
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
         Left            =   4860
         TabIndex        =   7
         Top             =   720
         Value           =   1  'Checked
         Width           =   2265
      End
      Begin VB.CheckBox chkEmpacador 
         Caption         =   "Todos los Empacadores"
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
         Left            =   4860
         TabIndex        =   5
         Top             =   330
         Value           =   1  'Checked
         Width           =   2355
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1590
         TabIndex        =   11
         Top             =   1530
         Width           =   1905
         _ExtentX        =   3360
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
         Format          =   63963137
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   1590
         TabIndex        =   8
         Top             =   1110
         Width           =   1905
         _ExtentX        =   3360
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
         Format          =   63963137
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   4860
         TabIndex        =   10
         Top             =   1140
         Width           =   1635
         _ExtentX        =   2884
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
         CustomFormat    =   "HH:mm"
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   4860
         TabIndex        =   12
         Top             =   1530
         Width           =   1635
         _ExtentX        =   2884
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
         CustomFormat    =   "HH:mm"
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboEmpacador 
         Height          =   315
         Left            =   1590
         TabIndex        =   4
         Top             =   330
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo cboMotorizado 
         Height          =   315
         Left            =   1590
         TabIndex        =   6
         Top             =   720
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
         Caption         =   "Formato :"
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
         Left            =   735
         TabIndex        =   19
         Top             =   1935
         Width           =   810
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motorizado :"
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
         Left            =   495
         TabIndex        =   18
         Top             =   780
         Width           =   1065
      End
      Begin VB.Label Label2 
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
         Left            =   330
         TabIndex        =   17
         Top             =   1215
         Width           =   1230
      End
      Begin VB.Label Label3 
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
         Left            =   435
         TabIndex        =   16
         Top             =   1620
         Width           =   1125
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empacador :"
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
         Index           =   8
         Left            =   480
         TabIndex        =   15
         Top             =   375
         Width           =   1080
      End
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
      Index           =   3
      Left            =   4500
      Picture         =   "frmRepTiempoDelivery.frx":0B62
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2460
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   750
      Top             =   2130
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepTiempoDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsReporte       As ADODB.Recordset
Dim RsEmpacador     As ADODB.Recordset
Dim RsMotorizado    As ADODB.Recordset
Dim Resumido       As New dsrDeliveryTiempo
Dim ResumidoCentral As New dsrDeliveryTiempoCD
Dim sCriterio       As String
Dim sTitulo         As String
Dim sNombre         As String

Dim fInicio As Date
Dim fFinal As Date

Sub LlenaCombos()
    With cboEmpacador
         Isql = "Select * from vEmpacador"
         Set RsEmpacador = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsEmpacador
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    
    With cboMotorizado
         Isql = "Select * from vMotorizado"
         Set RsMotorizado = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMotorizado
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
End Sub

Private Sub chkEmpacador_Click()
   If chkEmpacador.Value = 1 Then
      cboEmpacador.Enabled = False
      cboEmpacador.Text = ""
      RsEmpacador.Filter = adFilterNone
   Else
      cboEmpacador.Enabled = True
   End If
End Sub


Private Sub chkMotorizado_Click()
   If chkMotorizado.Value = 1 Then
      cboMotorizado.Enabled = False
      cboMotorizado.Text = ""
      RsMotorizado.Filter = adFilterNone
   Else
      cboMotorizado.Enabled = True
   End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   
   sCriterio = ""
   sTitulo = ""
   
      If dtpFecIni.Value + dtpHorIni.Value > dtpFecFin.Value + dtpHorFin.Value Then
         MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
         Exit Sub
         dtpFecFin.SetFocus
      End If
      sTitulo = " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.Value, "HH:mm") & " Al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.Value, "HH:mm")
       
   If chkEmpacador.Value = 0 Then
      If cboEmpacador.Text = "" Then
         MsgBox "Debe escoger un Empacador", vbCritical, sMensaje
         Exit Sub
      End If
    End If
   
   If chkMotorizado.Value = 0 Then
      If cboMotorizado.Text = "" Then
         MsgBox "Debe escoger un Motorizado", vbCritical, sMensaje
         Exit Sub
      End If
   End If
    
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               If optToma.Value = True Then
                  frmEmite.CRViewer.DisplayGroupTree = False
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crLandscape
                  frmEmite.CRViewer.ViewReport
                  frmEmite.Show vbModal
                Else
                  frmEmite.CRViewer.DisplayGroupTree = False
                  ResumidoCentral.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ResumidoCentral.PaperOrientation = crLandscape
                  frmEmite.CRViewer.ViewReport
                  frmEmite.Show vbModal
                End If
          
          Case Is = 1 ' Imprimir
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                   MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                   Exit Sub
               End If
               If optToma.Value = True Then
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crPortrait
                  Resumido.PrintOut
                Else
                  ResumidoCentral.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ResumidoCentral.PaperOrientation = crPortrait
                  ResumidoCentral.PrintOut
                End If
          Case Is = 2 ' Salir
               Unload Me
               
          Case Is = 3 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               If optToma.Value = True Then
                  Resumido.ExportOptions.FormatType = 21
                  Resumido.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  Resumido.ExportOptions.DiskFileName = cmdSave.FileName
                  Resumido.Export False
               Else
                  ResumidoCentral.ExportOptions.FormatType = 21
                  ResumidoCentral.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  ResumidoCentral.ExportOptions.DiskFileName = cmdSave.FileName
                  ResumidoCentral.Export False
               End If
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
    sNombre = ""
    sNombre = dbTemporal(sCaja, 16, "fFecha", "DATETIME", "tCodigoPedido", "NVARCHAR(10)", "tMesa", "NVARCHAR(3)", "nAdulto", "SMALLINT", "nMonto", "FLOAT", "nWL", "INT", "nUbicacion", "INT", "nPedido", "INT", "nSPlato", "INT", "nTiempoComer", "INT", "nIPedido", "INT", "nPago", "INT", "tTurno", "NVARCHAR(10)", "Salon", "NVARCHAR(50)", "tLocal", "NVARCHAR(10)", "tSalon", "NVARCHAR(2)")
    dtpFecIni.Value = Date
    dtpFecFin.Value = Date
    cboEmpacador.Enabled = False
    cboMotorizado.Enabled = False
    cboEmpacador.Text = ""
    cboMotorizado.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
   Set RsEmpacador = Nothing
   Set RsMotorizado = Nothing
   Cn.Execute "Drop Table " & sNombre
End Sub

Public Sub Genera()
   Dim oComando As clsComando
   
   Screen.MousePointer = vbHourglass
   Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_TiempoDelivery", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    fInicio = Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm")
    fFinal = Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm")
    oComando.CreateParameter "@tEmpacador", adVarChar, adParamInput, 20, Me.cboEmpacador.BoundText
    oComando.CreateParameter "@tMotorizado", adVarChar, adParamInput, 20, Me.cboMotorizado.BoundText
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@Tipo", adVarChar, adParamInput, 1, IIf(optToma.Value = True, "T", "C")
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set rsReporte = oComando.GetSP()
    If optToma.Value = True Then
            Resumido.DiscardSavedData
            Resumido.Database.SetDataSource rsReporte
            Resumido.Text14.SetText sRazonSocial
            Resumido.Text29.SetText localConectado
            Resumido.ReportTitle = sTitulo
            frmEmite.CRViewer.ReportSource = Resumido
    End If
    If optToma.Value = False Then
            ResumidoCentral.DiscardSavedData
            ResumidoCentral.Database.SetDataSource rsReporte
            ResumidoCentral.Text14.SetText sRazonSocial
            ResumidoCentral.Text29.SetText localConectado
            ResumidoCentral.ReportTitle = sTitulo & " Formato Central Delivery"
            frmEmite.CRViewer.ReportSource = ResumidoCentral
    
    End If
End Sub
