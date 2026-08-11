VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRepPrincipal 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Principales Clientes"
   ClientHeight    =   4155
   ClientLeft      =   2895
   ClientTop       =   2190
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepPrincipal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   8400
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
      Left            =   5490
      Picture         =   "frmRepPrincipal.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3510
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
      Height          =   3420
      Left            =   30
      TabIndex        =   14
      Top             =   30
      Width           =   8340
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   1515
         Width           =   2775
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
         Height          =   1275
         Left            =   1980
         TabIndex        =   19
         Top             =   2010
         Width           =   3615
         Begin VB.OptionButton optOpcion 
            Caption         =   "Resumido"
            Height          =   240
            Index           =   1
            Left            =   210
            TabIndex        =   9
            Top             =   780
            Width           =   2835
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Detallado"
            Height          =   240
            Index           =   0
            Left            =   210
            TabIndex        =   8
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
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1095
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
         Left            =   4875
         Picture         =   "frmRepPrincipal.frx":082E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1065
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
         Left            =   5880
         TabIndex        =   6
         Top             =   1110
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   7170
         Top             =   330
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "Xls"
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   330
         Left            =   1995
         TabIndex        =   2
         Top             =   660
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         _Version        =   393216
         Format          =   63963137
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   330
         Left            =   1995
         TabIndex        =   0
         Top             =   225
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         _Version        =   393216
         Format          =   63963137
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   330
         Left            =   4035
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
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   330
         Left            =   4035
         TabIndex        =   3
         Top             =   660
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
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
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
         Left            =   1170
         TabIndex        =   18
         Top             =   1155
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Mínimo :"
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
         Left            =   555
         TabIndex        =   17
         Top             =   1575
         Width           =   1335
      End
      Begin VB.Label Label 
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
         Left            =   765
         TabIndex        =   16
         Top             =   728
         Width           =   1125
      End
      Begin VB.Label Label 
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
         Left            =   660
         TabIndex        =   15
         Top             =   293
         Width           =   1230
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
      Left            =   6945
      Picture         =   "frmRepPrincipal.frx":0930
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3510
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
      Left            =   4035
      Picture         =   "frmRepPrincipal.frx":0A22
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3510
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
      Left            =   2580
      Picture         =   "frmRepPrincipal.frx":0F54
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3510
      Width           =   1455
   End
End
Attribute VB_Name = "frmRepPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCliente As Recordset
Dim Resumido As New dsrPrincipalR
Dim Detallado As New dsrPrincipalD
Dim RsPrinter As Recordset
Dim sCliente As String
Dim Identidad As String
Dim fInicio As Date
Dim fFinal As Date

Private Sub cmdBusca_Click()
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
   'sCliente = Calcular("Select tCodigoCliente As Codigo From tCliente Where tIdentidad ='" & sCodigo & "'", Cn)
   sCliente = sCodigo
   txtCliente.Text = sDescrip
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 3 Then
      Unload Me
      Exit Sub
   End If
   

   If dtpFecIni.Value + dtpHorIni.Value > dtpFecFin.Value + dtpHorFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      Exit Sub
      dtpFecFin.SetFocus
   End If
      
   If chkCliente.Value = 0 Then
      If sCliente = "" Then
         MsgBox "Debe escoger el Cliente", vbCritical, sMensaje
         Exit Sub
      End If
   End If
   fInicio = Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm")
    fFinal = Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm")
      
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If RsPrinter.EOF = True Then
                  Screen.MousePointer = vbDefault
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
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
               If RsPrinter.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               If optOpcion(0).Value = True Then
                  Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Detallado.PaperOrientation = crPortrait
                  Detallado.PrintOut
               Else
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crPortrait
                  Resumido.PrintOut
               End If
          
          Case Is = 2 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If RsPrinter.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
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
                     Exit Sub
                  End If
                  Resumido.ExportOptions.DiskFileName = cmdSave.FileName
                  Resumido.Export False
               End If
         
          Case Is = 3 ' Salir
               Unload Me
   End Select
   
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

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.Value > dtpFecFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub Form_Load()
   Centrar Me
   dtpFecIni.Value = Date
   dtpFecFin.Value = Date
   cmdBusca.Enabled = False
End Sub
Public Sub Genera()
   Dim sOrden1 As String
   Dim sOrden2 As String
   Dim wCodigo As String
   Dim oComando As clsComando
    
   Screen.MousePointer = vbHourglass
       Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_PrincipalCliente", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    
    oComando.CreateParameter "@flagTipo", adBoolean, adParamInput, 1, optOpcion(0).Value
    oComando.CreateParameter "@sMonto", adDouble, adParamInput, 10, CDbl(txtMonto.Text)
    oComando.CreateParameter "@sCliente", adVarChar, adParamInput, 10, sCliente
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set RsPrinter = oComando.GetSP()
    
   If optOpcion(0).Value = True Then
      Detallado.DiscardSavedData
      Detallado.Database.SetDataSource RsPrinter
      Detallado.ReportTitle = "Del " & dtpFecIni.Value & " Al " & dtpFecFin.Value
      Detallado.Text3.SetText sImpuesto1
      Detallado.Text10.SetText sImpuesto2
      Detallado.Text11.SetText sImpuesto3
      Detallado.Text1.SetText localConectado
      Detallado.Text28.SetText sRazonSocial
      frmEmite.CRViewer.DisplayGroupTree = False
      frmEmite.CRViewer.ReportSource = Detallado
   Else
      Resumido.DiscardSavedData
      Resumido.Database.SetDataSource RsPrinter
      Resumido.ReportTitle = "Del " & dtpFecIni.Value & " Al " & dtpFecFin.Value
      Resumido.Text3.SetText sImpuesto1
      Resumido.Text10.SetText sImpuesto2
      Resumido.Text11.SetText sImpuesto3
      Resumido.Text1.SetText localConectado
      Resumido.Text28.SetText sRazonSocial
      frmEmite.CRViewer.DisplayGroupTree = False
      frmEmite.CRViewer.ReportSource = Resumido
   End If
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsPrinter = Nothing
   Set frmRepPrincipal = Nothing
   Set RsCliente = Nothing
End Sub

Private Sub txtMonto_Change()
   If txtMonto.Text = "" Or Not IsNumeric(txtMonto) Then
      txtMonto.Text = "0.00"
   End If
End Sub

Private Sub txtMonto_gotfocus()
   txtMonto.SelLength = Len(Trim(txtMonto.Text))
End Sub

Private Sub txtMonto_LostFocus()
   txtMonto.Text = Format(txtMonto.Text, "###,##0.00")
End Sub

