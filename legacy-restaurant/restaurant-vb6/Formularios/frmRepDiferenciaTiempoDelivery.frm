VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepDiferenciaTiempoDelivery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Diferencias de Tiempos Delivery"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "frmRepDiferenciaTiempoDelivery.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
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
      Picture         =   "frmRepDiferenciaTiempoDelivery.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4260
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
      Picture         =   "frmRepDiferenciaTiempoDelivery.frx":053E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4260
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
      Picture         =   "frmRepDiferenciaTiempoDelivery.frx":0A70
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4260
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
      Height          =   4155
      Left            =   0
      TabIndex        =   1
      Top             =   45
      Width           =   7410
      Begin VB.TextBox txtFinal 
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
         Left            =   5535
         MaxLength       =   3
         TabIndex        =   25
         Text            =   "0"
         Top             =   2595
         Width           =   555
      End
      Begin VB.TextBox txtInicio 
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
         Left            =   5535
         MaxLength       =   3
         TabIndex        =   24
         Text            =   "0"
         Top             =   2100
         Width           =   555
      End
      Begin VB.Frame Frame2 
         Caption         =   " Tiempos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   1590
         TabIndex        =   17
         Top             =   2025
         Width           =   3135
         Begin VB.CheckBox chkTiempo 
            Caption         =   "Asignación del Motorizado"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   22
            Top             =   966
            Width           =   2355
         End
         Begin VB.CheckBox chkTiempo 
            Caption         =   "Fecha de Llegada"
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   21
            Top             =   1575
            Width           =   2355
         End
         Begin VB.CheckBox chkTiempo 
            Caption         =   "Fecha de Salida"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   20
            Top             =   1269
            Width           =   2355
         End
         Begin VB.CheckBox chkTiempo 
            Caption         =   "Asignación del Empacador"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   19
            Top             =   663
            Width           =   2355
         End
         Begin VB.CheckBox chkTiempo 
            Caption         =   "Fecha del Pedido"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   18
            Top             =   360
            Width           =   2355
         End
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   330
         Value           =   1  'Checked
         Width           =   2355
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1590
         TabIndex        =   4
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
         Format          =   56295425
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   1590
         TabIndex        =   5
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
         Format          =   56295425
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   4860
         TabIndex        =   6
         Top             =   1125
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
         Format          =   56295427
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   4860
         TabIndex        =   7
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
         Format          =   56295427
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboEmpacador 
         Height          =   315
         Left            =   1590
         TabIndex        =   8
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
         TabIndex        =   9
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
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
         Left            =   4875
         TabIndex        =   29
         Top             =   2655
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "desde"
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
         Left            =   4860
         TabIndex        =   28
         Top             =   2160
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "min"
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
         Left            =   6195
         TabIndex        =   27
         Top             =   2655
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "min"
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
         Left            =   6195
         TabIndex        =   26
         Top             =   2160
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diferencias :"
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
         Left            =   465
         TabIndex        =   23
         Top             =   2025
         Width           =   1095
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
      Picture         =   "frmRepDiferenciaTiempoDelivery.frx":0B62
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4260
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   990
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepDiferenciaTiempoDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsReporte       As ADODB.Recordset
Dim RsEmpacador     As ADODB.Recordset
Dim RsMotorizado    As ADODB.Recordset
Dim Resumido       As New dsrDiferenciaTiempo
Dim sCriterio       As String
Dim sTitulo         As String
Dim sNombre         As String

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

Private Sub chkTiempo_Click(Index As Integer)
  Dim i As Integer
  Dim nTiempo As Integer
  nTiempo = 0
  For i = 0 To 4
      If chkTiempo(i).Value Then
         nTiempo = nTiempo + 1
      End If
  Next i
  If nTiempo > 2 Then
     chkTiempo(Index).Value = False
  End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   
   sCriterio = ""
   sTitulo = ""
        
   Dim nCont As Integer
   Dim i As Integer
   For i = 0 To 4
       If chkTiempo(i).Value Then
          nCont = nCont + 1
       End If
   Next i
   
   If nCont <> 2 Then
      MsgBox "Error, Marque dos opciones de tiempo", , sMensaje
      Exit Sub
   End If
   
   If Not IsNumeric(txtInicio.Text) Then
      MsgBox "Error, Ingrese un valor numérico", , sMensaje
      txtInicio.SetFocus
      Exit Sub
   End If
      
   If Not IsNumeric(txtFinal.Text) Then
      MsgBox "Error, Ingrese un valor numérico", , sMensaje
      txtFinal.SetFocus
      Exit Sub
   End If
      
   If Not (Val(txtInicio.Text) > 0 And Val(txtFinal.Text) > 0 And Val(txtFinal.Text) - Val(txtInicio.Text) > 0) Then
      MsgBox "Error en rango de minutos", vbCritical, sMensaje
      Exit Sub
   End If
      
   If dtpFecIni.Value + dtpHorIni.Value > dtpFecFin.Value + dtpHorFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      Exit Sub
      dtpFecFin.SetFocus
   End If
   sTitulo = " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.Value, "HH:mm") & " Al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.Value, "HH:mm")
   sCriterio = " AND MPEDIDO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm") & "' and MPEDIDO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm") & "' "
   
   If chkEmpacador.Value = 0 Then
      If cboEmpacador.Text = "" Then
         MsgBox "Debe escoger un Empacador", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " and mpedido.tempacador ='" & cboEmpacador.BoundText & "'"
   End If
   
   If chkMotorizado.Value = 0 Then
      If cboMotorizado.Text = "" Then
         MsgBox "Debe escoger un Motorizado", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & "and MPEDIDO.tmotorizado ='" & cboMotorizado.BoundText & "'"
   End If
    
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
                  frmEmite.CRViewer.DisplayGroupTree = False
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crLandscape
                  frmEmite.CRViewer.ViewReport
                  frmEmite.Show vbModal
          
          Case Is = 1 ' Imprimir
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                   MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                   Exit Sub
               End If
                  Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Resumido.PaperOrientation = crPortrait
                  Resumido.PrintOut
                 
          Case Is = 2 ' Salir
               Unload Me
               
          Case Is = 3 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
                  Resumido.ExportOptions.FormatType = 21
                  Resumido.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  Resumido.ExportOptions.DiskFileName = cmdSave.FileName
                  Resumido.Export False
       
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
   Dim sCampo As String
   Dim sListado As String
   Dim lPrimera As Boolean
   lPrimera = True
   
   sCampo = ""
   sListado = ""
   sTitulo = sTitulo & " Diferencia de "
   If chkTiempo(4).Value Then
      lPrimera = False
      sCampo = ",MPEDIDO.fLlegada"
      sListado = ",convert(nvarchar(16),MPEDIDO.fLlegada,120) as fEmpacador"
      sTitulo = sTitulo & " Fecha de Llegada y"
      Resumido.Text19.SetText "Fecha de Llegada"
   End If
       
   If chkTiempo(3).Value Then
      sCampo = sCampo & ",MPEDIDO.fSalida"
      
      sTitulo = sTitulo & " Fecha de Salida"
      If lPrimera Then
        sListado = sListado & ",convert(nvarchar(16),MPEDIDO.fSalida,120) as fEmpacador"
         lPrimera = False
         sTitulo = sTitulo & " y"
         Resumido.Text19.SetText "Fecha de Salida"
      Else
        sListado = sListado & ",convert(nvarchar(16),MPEDIDO.fSalida,120) as fRegistro"
         Resumido.Text20.SetText "Fecha de Salida"
      End If
   End If
       
   If chkTiempo(2).Value Then
      sCampo = sCampo & ",MPEDIDO.fAsignacion"
      
      sTitulo = sTitulo & " Fecha del Motorizado"
      If lPrimera Then
        sListado = sListado & ",convert(nvarchar(16),MPEDIDO.fAsignacion,120) as fEmpacador"
         lPrimera = False
         sTitulo = sTitulo & " y"
         Resumido.Text19.SetText "Fecha del Motorizado"
      Else
        sListado = sListado & ",convert(nvarchar(16),MPEDIDO.fAsignacion,120) as fRegistro"
         Resumido.Text20.SetText "Fecha del Motorizado"
      End If
   End If
       
   If chkTiempo(1).Value Then
      sCampo = sCampo & ",MPEDIDO.fEmpacador"
     
      sTitulo = sTitulo & " Fecha de Empacador"
      If lPrimera Then
         sListado = sListado & ",convert(nvarchar(16),MPEDIDO.fEmpacador,120) as fEmpacador"
         lPrimera = False
         sTitulo = sTitulo & " y"
         Resumido.Text19.SetText "Fecha del Empacador"
      Else
         sListado = sListado & ",convert(nvarchar(16),MPEDIDO.fEmpacador,120) as fRegistro"
         Resumido.Text20.SetText "Fecha del Empacador"
      End If
   End If
   
   If chkTiempo(0).Value Then
      sCampo = sCampo & ",MPEDIDO.fRegistro"
      sListado = sListado & ",convert(nvarchar(16),MPEDIDO.fRegistro,120) as fRegistro"
      sTitulo = sTitulo & " Fecha de Registro"
      Resumido.Text20.SetText "Fecha de Registro"
   End If
   sTitulo = sTitulo & " mayor a " & Trim(txtInicio.Text) & " y menor a " & Trim(txtFinal.Text) & " minutos"
   Screen.MousePointer = vbHourglass
   
   sCriterio = sCriterio & " and isnull(datediff(minute" & sCampo & "),0) * -1 >= " & Trim(txtInicio.Text) & " and isnull(datediff(minute" & sCampo & "),0) * -1 <= " & Trim(txtFinal.Text)

   Isql = "SELECT dbo.MPEDIDO.tCodigoPedido AS pedido, isnull(dbo.vZona.Descripcion,'') AS Zona, dbo.TDELIVERY.tDireccion AS Direccion, LTRIM(dbo.TDELIVERY.tApellido) + ' ' + LTRIM(dbo.TDELIVERY.tNombre) AS Nombre, dbo.TDELIVERY.tTelefono, SUM(dbo.DPEDIDO.nVenta) AS nVenta " & sListado & ", " & _
          "ISNULL(DATEDIFF(minute " & sCampo & "), 0) * -1 AS nDiferencia, dbo.vMotorizado.Descripcion AS Cliente " & _
          "FROM dbo.MPEDIDO INNER JOIN dbo.TDELIVERY ON dbo.MPEDIDO.tClienteDelivery = dbo.TDELIVERY.tCodigoDelivery INNER JOIN dbo.DPEDIDO ON dbo.MPEDIDO.tCodigoPedido = dbo.DPEDIDO.tCodigoPedido LEFT OUTER JOIN dbo.vMotorizado ON dbo.MPEDIDO.tMotorizado = dbo.vMotorizado.Codigo LEFT OUTER JOIN dbo.vZona ON dbo.TDELIVERY.tZona = dbo.vZona.Codigo " & _
          "where mpedido.testadopedido<>'03' and mpedido.tTipoPedido='02'" & sCriterio & _
          "GROUP BY dbo.MPEDIDO.tCodigoPedido, dbo.vZONA.Descripcion, dbo.TDELIVERY.tDireccion, LTRIM(dbo.TDELIVERY.tApellido) + ' ' + LTRIM(dbo.TDELIVERY.tNombre), dbo.TDELIVERY.tTelefono " & sCampo & ", isnull(datediff(minute" & sCampo & "),0) * - 1, dbo.vMotorizado.Descripcion " & _
          "ORDER BY dbo.MPEDIDO.tCodigoPedido"
   Set rsReporte = Lib.OpenRecordset(Isql, Cn)
      
   Resumido.DiscardSavedData
   Resumido.Database.SetDataSource rsReporte
   Resumido.Text14.SetText sRazonSocial
   Resumido.Text7.SetText localConectado
   Resumido.ReportTitle = sTitulo
   frmEmite.CRViewer.ReportSource = Resumido
End Sub
