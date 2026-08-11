VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepDocumentosAnticipos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  Reporte de Documentos"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
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
      Height          =   3780
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8985
      Begin VB.CheckBox chkTodosCFact 
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
         Height          =   255
         Left            =   7920
         TabIndex        =   20
         Top             =   960
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkTodosCFrec 
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
         Height          =   255
         Left            =   7920
         TabIndex        =   19
         Top             =   480
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.OptionButton optDocEmitAnt 
         Caption         =   " Documentos Emitidos por anticipo"
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
         Left            =   1920
         TabIndex        =   18
         Top             =   2640
         Width           =   3375
      End
      Begin VB.OptionButton optConAnt 
         Caption         =   " Consumo de Anticipos"
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
         Left            =   1920
         TabIndex        =   17
         Top             =   3000
         Width           =   2295
      End
      Begin VB.OptionButton optTodos 
         Caption         =   " Todos"
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
         Left            =   1920
         TabIndex        =   16
         Top             =   3360
         Value           =   -1  'True
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   1920
         Width           =   1905
         _ExtentX        =   3360
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
         Format          =   63963137
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   1440
         Width           =   1905
         _ExtentX        =   3360
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
         Format          =   63963137
         CurrentDate     =   37539.2083333333
      End
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   6240
         Top             =   1680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo cboCaja 
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Top             =   480
         Width           =   3345
         _ExtentX        =   5900
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
      Begin MSComCtl2.DTPicker dtpHoraIni 
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   1440
         Width           =   1350
         _ExtentX        =   2381
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
      Begin MSComCtl2.DTPicker dtpHoraFin 
         Height          =   375
         Left            =   3960
         TabIndex        =   9
         Top             =   1920
         Width           =   1350
         _ExtentX        =   2381
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
      Begin MSDataListLib.DataCombo cboCFact 
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   960
         Width           =   5865
         _ExtentX        =   10345
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar solo:"
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
         Left            =   240
         TabIndex        =   15
         Top             =   2640
         Width           =   960
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente Facturado:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Caja:"
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
         Left            =   1360
         TabIndex        =   13
         Top             =   480
         Width           =   495
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
         Left            =   240
         TabIndex        =   11
         Top             =   1440
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
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   1125
      End
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
      Left            =   1725
      Picture         =   "frmDocumentosAnticipos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4110
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
      Left            =   3180
      Picture         =   "frmDocumentosAnticipos.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4110
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
      Index           =   3
      Left            =   6105
      Picture         =   "frmDocumentosAnticipos.frx":0A64
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4110
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
      Left            =   4635
      Picture         =   "frmDocumentosAnticipos.frx":0B56
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4110
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmRepDocumentosAnticipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsReporte As Recordset
Dim RsUser    As Recordset
Dim RsMozo    As Recordset
Dim sTitulo   As String
Dim sCriterio As String
Dim sOrden    As String
Dim RsEmpresa As Recordset
Dim RsDocumento As Recordset

Dim CrtAnticipoConsumos As New dsrAnticipoAmbosConsumos
Dim CrtAnticipo As New dsrAnticipoAmbos

Dim sReporte As String

Dim fInicio As Date
Dim fFinal As Date

Private Sub chkTodosCFact_Click()
   cboCFact.Enabled = Not cboCFact.Enabled
   cboCFact.BoundText = ""
End Sub

Private Sub chkTodosCFrec_Click()
   cboCaja.Enabled = Not cboCaja.Enabled
   cboCaja.BoundText = ""
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
 sCriterio = ""
 Screen.MousePointer = vbHourglass

   If chkTodosCFrec.Value = 0 And cboCaja.BoundText = "" And chkTodosCFact.Value = 0 And cboCFact.BoundText = "" Then
      MsgBox "Seleccione Cliente Frecuente o Cliente Facturado", vbExclamation, "Mensaje del Sistema"
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   
   If chkTodosCFrec.Value = 0 And cboCaja.BoundText = "" Then
      MsgBox "Seleccione Cliente Frecuente", vbExclamation, "Mensaje del Sistema"
      cboCFrec.SetFocus
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   
   If chkTodosCFact.Value = 0 And cboCFact.BoundText = "" Then
      MsgBox "Seleccione Cliente Facturado", vbExclamation, "Mensaje del Sistema"
      cboCFact.BoundText = ""
      cboCFact.SetFocus
      Screen.MousePointer = vbDefault
      Exit Sub
   End If

   If Index = 3 Then
      Unload Me
   Else
      Genera Index
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    dtpFecIni.Value = FechaServidor
    dtpFecFin.Value = FechaServidor

    LlenaCombos
End Sub

Private Sub LlenaCombos()
   With cboCaja
      Isql = "select * from tcaja "
      Set RsTipo = Lib.OpenRecordset(Isql, Cn)
      Set .RowSource = RsTipo
      .DataField = "tDescripcion"
      .ListField = "tDescripcion"
      .BoundColumn = "tCaja"
   End With

   With cboCFact
      Isql = "select upper(ltrim(rtrim(codigo))) CODIGO, upper(ltrim(rtrim(descripcion))) CLIENTE from vcliente order by 2"
      Set RsTipo = Lib.OpenRecordset(Isql, Cn)
      Set .RowSource = RsTipo
      .DataField = "CLIENTE"
      .ListField = "CLIENTE"
      .BoundColumn = "CODIGO"
   End With
End Sub

Public Sub Genera(iOpcion As Integer)
On Error GoTo fin
   Dim oComando As clsComando

   Screen.MousePointer = vbHourglass
   Set oComando = New clsComando
   If Not oComando.CreateCmdSp("usp_Anticipos", Cn) Then
      Set oComando = Nothing
      Exit Sub
   End If

   fInicio = Format(dtpFecIni.Value, "yyyy/MM/dd") & " " & Format(dtpHoraIni.Value, "HH:mm")
   fFinal = Format(dtpFecFin.Value, "yyyy/MM/dd") & " " & Format(dtpHoraFin.Value, "HH:mm")

'   If chkTodosCFrec.value = 1 And chkTodosCFact.value = 1 And (optTodos Or optDocEmitAnt) And Not optConAnt Then
'      oComando.CreateParameter "@opcion", adInteger, adParamInput, 1, 1
'      oComando.CreateParameter "@codigocliente", adVarChar, adParamInput, 15, Me.cboCFact.BoundText
'   End If
'
'   If chkTodosCFrec.value = 0 And chkTodosCFact.value = 1 And (optTodos Or optDocEmitAnt) And Not optConAnt Then
'      oComando.CreateParameter "@opcion", adInteger, adParamInput, 1, 2
'      oComando.CreateParameter "@codigocliente", adVarChar, adParamInput, 15, cboCFrec.BoundText
'   End If
'
'   If chkTodosCFrec.value = 1 And chkTodosCFact.value = 0 And (optTodos Or optDocEmitAnt) And Not optConAnt Then
'      oComando.CreateParameter "@opcion", adInteger, adParamInput, 1, 3
'      oComando.CreateParameter "@codigocliente", adVarChar, adParamInput, 15, cboCFact.BoundText
'   End If
'
'   If chkTodosCFrec.value = 0 And chkTodosCFact.value = 1 And Not optTodos And Not optDocEmitAnt And optConAnt Then
'      oComando.CreateParameter "@opcion", adInteger, adParamInput, 1, 4
'      oComando.CreateParameter "@codigocliente", adVarChar, adParamInput, 15, cboCFrec.BoundText
'   End If
'
'   If chkTodosCFrec.value = 1 And chkTodosCFact.value = 1 And Not optTodos And Not optDocEmitAnt And optConAnt Then
'      oComando.CreateParameter "@opcion", adInteger, adParamInput, 1, 5
'      oComando.CreateParameter "@codigocliente", adVarChar, adParamInput, 15, ""
'   End If

    oComando.CreateParameter "@opcion", adInteger, adParamInput, 1, IIf(optTodos, 1, IIf(optDocEmitAnt, 2, IIf(Me.optConAnt, 3, 1)))
    oComando.CreateParameter "@codigocliente", adVarChar, adParamInput, 15, Me.cboCFact.BoundText
    oComando.CreateParameter "@codigoFrecuente", adVarChar, adParamInput, 15, Me.cboCaja.BoundText
    oComando.CreateParameter "@desde", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@hasta", adDBDate, adParamInput, 10, fFinal

   If Not oComando.GetParamOK Then
      Set oComando = Nothing
      Exit Sub
   End If

   Set rsReporte = oComando.GetSP()
   'rsReporte.Filter = sCriterio

   If rsReporte.EOF = True Then
      MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   
   'If optTodos Or optDocEmitAnt Then
      Dim CrtAnticipo As New dsrAnticipoAmbos
      
      CrtAnticipo.DiscardSavedData
      CrtAnticipo.Database.SetDataSource rsReporte

      CrtAnticipo.Text17.SetText sRazonSocial
    
      If optTodos Then
        CrtAnticipo.txtSubTitulo.SetText "Del: " & fInicio & " Al:" & fFinal & " <<Todos>>"
        'CrtAnticipo.Field2.TextColor = 1
      End If
      
      If optDocEmitAnt Then
        CrtAnticipo.txtSubTitulo.SetText "Del: " & fInicio & " Al:" & fFinal & " <<Anticipo>>"
        'CrtAnticipo.Field2.TextColor = &HFFFF&
      End If
      
      If optConAnt Then
        CrtAnticipo.txtSubTitulo.SetText "Del: " & fInicio & " Al:" & fFinal & " <<Consumo Anticipo>>"
        'CrtAnticipo.Field13.TextColor = &HFFFFFF
        CrtAnticipo.Text4.SetText "TOTAL CONSUMO:"
        
      End If
      
      frmEmite.CRViewer.ReportSource = CrtAnticipo
      frmEmite.CRViewer.DisplayGroupTree = False
      CrtAnticipo.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
      CrtAnticipo.PaperOrientation = crPortrait

      If iOpcion = 1 Then CrtAnticipo.PrintOut
      
      If iOpcion = 2 Then
         CrtAnticipo.ExportOptions.FormatType = 21
         CrtAnticipo.ExportOptions.DestinationType = 1
         cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
         cmdSave.ShowSave
         If cmdSave.FileName = "" Then Exit Sub
         
         CrtAnticipo.ExportOptions.DiskFileName = cmdSave.FileName
         CrtAnticipo.Export False
      End If

'   Else
'      Dim CrtAnticipoConsumos As New dsrAnticipoAmbosConsumos
'
'      CrtAnticipoConsumos.DiscardSavedData
'      CrtAnticipoConsumos.Database.SetDataSource rsReporte
'      'CrtAnticipo.ReportTitle = sTitulo
'
'      frmEmite.CRViewer.ReportSource = CrtAnticipoConsumos
'      frmEmite.CRViewer.DisplayGroupTree = False
'      CrtAnticipoConsumos.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
'      CrtAnticipoConsumos.PaperOrientation = crPortrait
'
'      If iOpcion = 1 Then CrtAnticipoConsumos.PrintOut
'
'      If iOpcion = 2 Then
'         CrtAnticipoConsumos.ExportOptions.FormatType = 21
'         CrtAnticipoConsumos.ExportOptions.DestinationType = 1
'         cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
'         cmdSave.ShowSave
'         If cmdSave.FileName = "" Then Exit Sub
'
'         CrtAnticipoConsumos.ExportOptions.DiskFileName = cmdSave.FileName
'         CrtAnticipoConsumos.Export False
'      End If
'
'   End If

   If iOpcion = 0 Then
      frmEmite.CRViewer.ViewReport
      frmEmite.Show vbModal
   End If
Exit Sub
fin:
    MsgBox "Error" & error
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
End Sub
