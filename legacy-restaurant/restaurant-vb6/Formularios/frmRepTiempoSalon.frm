VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepTiempoSalon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tiempos en Salon"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "frmRepTiempoSalon.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
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
      Left            =   1635
      Picture         =   "frmRepTiempoSalon.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2595
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
      Left            =   3090
      Picture         =   "frmRepTiempoSalon.frx":053E
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2595
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
      Left            =   6000
      Picture         =   "frmRepTiempoSalon.frx":0A70
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2595
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
      Height          =   2535
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7455
      Begin VB.CheckBox chkSalon 
         Caption         =   "Todos los Salones"
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
         Left            =   5130
         TabIndex        =   6
         Top             =   736
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.CheckBox chkTurno 
         Caption         =   "Todos los Turnos"
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
         Left            =   5130
         TabIndex        =   5
         Top             =   1125
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkLocal 
         Caption         =   "Todos los Locales"
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
         Left            =   5130
         TabIndex        =   4
         Top             =   330
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.TextBox txtTurno 
         Enabled         =   0   'False
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
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   2265
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
         Left            =   4230
         Picture         =   "frmRepTiempoSalon.frx":0B62
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1050
         Width           =   765
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1860
         TabIndex        =   7
         Top             =   1935
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
         Left            =   1860
         TabIndex        =   8
         Top             =   1515
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
         Left            =   5130
         TabIndex        =   9
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
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   5130
         TabIndex        =   10
         Top             =   1935
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
      Begin MSDataListLib.DataCombo cboLocal 
         Height          =   315
         Left            =   1860
         TabIndex        =   11
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
      Begin MSDataListLib.DataCombo cboSalon 
         Height          =   315
         Left            =   1860
         TabIndex        =   12
         Top             =   726
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
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salón :"
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
         Left            =   1170
         TabIndex        =   17
         Top             =   781
         Width           =   615
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Turno :"
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
         TabIndex        =   16
         Top             =   1170
         Width           =   630
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
         Left            =   555
         TabIndex        =   15
         Top             =   1620
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
         Left            =   660
         TabIndex        =   14
         Top             =   2025
         Width           =   1125
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local :"
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
         Left            =   1185
         TabIndex        =   13
         Top             =   375
         Width           =   600
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
      Left            =   4545
      Picture         =   "frmRepTiempoSalon.frx":0C64
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2595
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   750
      Top             =   2625
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepTiempoSalon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsReporte       As ADODB.Recordset
Dim Rsturno         As ADODB.Recordset
Dim RsLocal         As ADODB.Recordset
Dim RsSalon         As ADODB.Recordset
Dim Detallado       As New dsrTiempoSalon
Dim sCriterio       As String
Dim sTitulo         As String
Dim fInicio         As Date
Dim fFinal          As Date

Sub LlenaCombos()
    With cboLocal
        Isql = "Select * from vLocal"
        Set RsLocal = Lib.OpenRecordset(Isql, Cn)
        Set .RowSource = RsLocal
        .DataField = "Descripcion"
        .ListField = "Descripcion"
        .BoundColumn = "Codigo"
    End With
    
    With cboSalon
        Isql = "Select * from vSalon"
        Set RsSalon = Lib.OpenRecordset(Isql, Cn)
        Set .RowSource = RsSalon
        .DataField = "Descripcion"
        .ListField = "Descripcion"
        .BoundColumn = "Codigo"
    End With
End Sub

Private Sub cboLocal_Click(Area As Integer)
    cboSalon.Text = ""
    With cboSalon
        Isql = "Select * from vSalon " & IIf(cboLocal.BoundText = "", "", "where tLocal = '" & cboLocal.BoundText & "'") & " order by Descripcion "
        Set RsSalon = Lib.OpenRecordset(Isql, Cn)
        Set .RowSource = RsSalon
    End With
End Sub

Private Sub cmdBusca_Click()
    Isql = "Select tTurno as Codigo, tCaja, convert(nVarChar,fInicial, 120) as Descripcion, convert(nvarchar,fFinal,120) as fFinal from MTURNO order by tTurno DESC"
    frmBusca.nPredeterm = 2
    Call ConfGrilla(4, frmBusca.grdGrilla, "Turno", 2, "Codigo", 1400, 0, 0, "", _
                                           "Caja", 2, "tCaja", 1100, 0, 0, "", _
                                           "Fecha Inicial", 2, "Descripcion", 2400, 0, 0, "", _
                                           "Fecha Final", 2, "fFinal", 2400, 0, 0, "")
       
    frmBusca.Show vbModal
    If Not wEnter Then
       Exit Sub
    End If
    sTurno = sCodigo
    txtTurno.Text = sCodigo
End Sub

Private Sub chkLocal_Click()
    If chkLocal.Value = 1 Then
        cboLocal.Enabled = False
        cboLocal.Text = ""
        RsSalon.Filter = adFilterNone
    Else
        cboLocal.Enabled = True
    End If
End Sub

Private Sub chkSalon_Click()
    If chkSalon.Value = 1 Then
        cboSalon.Enabled = False
        cboSalon.Text = ""
    Else
        cboSalon.Enabled = True
    End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    If Index = 2 Then
        Unload Me
        Exit Sub
    End If
   
    sCriterio = ""
    sTitulo = ""
   
    If chkTurno.Value = 0 Then
        If sTurno = "" Then
            MsgBox "Debe escoger un Turno", vbCritical, sMensaje
            Exit Sub
        End If
              
        Set Rsturno = Lib.OpenRecordset("select * from MTURNO where tTurno='" & sTurno & "'", Cn)
        If Rsturno.RecordCount > 0 Then
            sTitulo = " Del " & Format(Rsturno!finicial, "dd/MMM/yyyy HH:mm") & " Al " & Format(IIf(IsNull(Rsturno!fFinal), Format(Rsturno!finicial, "dd/mmm/yyyy") & " 23:59", Rsturno!fFinal), "dd/mmm/yyyy HH:mm")
        Else
            MsgBox "Error Fatal en Turnos", vbCritical, sMensaje
        End If
    Else
        If dtpFecIni.Value + dtpHorIni.Value > dtpFecFin.Value + dtpHorFin.Value Then
            MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
            Exit Sub
            dtpFecFin.SetFocus
        End If
        sTitulo = " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.Value, "HH:mm") & " Al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.Value, "HH:mm")
    End If
   
    If chkLocal.Value = 0 Then
        If cboLocal.Text = "" Then
            MsgBox "Debe escoger un Local", vbCritical, sMensaje
            Exit Sub
        End If
    End If
   
    If chkSalon.Value = 0 Then
        If cboSalon.Text = "" Then
            MsgBox "Debe escoger un Salón", vbCritical, sMensaje
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
            frmEmite.CRViewer.DisplayGroupTree = True
            Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
            Detallado.PaperOrientation = crLandscape
            frmEmite.CRViewer.ViewReport
            frmEmite.Show vbModal
          
        Case Is = 1 ' Imprimir
            Genera
            Screen.MousePointer = vbDefault
            If rsReporte.EOF = True Then
                MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                Exit Sub
            End If
            Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
            Detallado.PaperOrientation = crLandscape
            Detallado.PrintOut
          
        Case Is = 2 ' Salir
            Unload Me
               
        Case Is = 3 ' Exportar
            Genera
            Screen.MousePointer = vbDefault
            If rsReporte.EOF = True Then
                MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                Exit Sub
            End If
            Detallado.ExportOptions.FormatType = 21
            Detallado.ExportOptions.DestinationType = 1
            cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
            cmdSave.ShowSave
            If cmdSave.FileName = "" Then
               Exit Sub
            End If
            Detallado.ExportOptions.DiskFileName = cmdSave.FileName
            Detallado.Export False
   End Select
End Sub

Private Sub dtpFecfin_LostFocus()
    If dtpFecIni.Value > dtpFecFin.Value Then
        MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
        dtpFecFin.SetFocus
    End If
End Sub

Private Sub ChkTurno_Click()
    If chkTurno.Value = 1 Then
        cmdBusca.Enabled = False
        sTurno = ""
        dtpFecIni.Enabled = True
        dtpFecFin.Enabled = True
        dtpHorIni.Enabled = True
        dtpHorFin.Enabled = True
    Else
        cmdBusca.Enabled = True
        dtpFecIni.Enabled = False
        dtpFecFin.Enabled = False
        dtpHorIni.Enabled = False
        dtpHorFin.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Centrar Me
    LlenaCombos
   
    dtpFecIni.Value = Date
    dtpFecFin.Value = Date
    cmdBusca.Enabled = False
    sTurno = ""
    cboLocal.Enabled = False
    cboSalon.Enabled = False
    cboLocal.Text = ""
    cboSalon.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsReporte = Nothing
    Set Rsturno = Nothing
    Set RsLocal = Nothing
    Set RsSalon = Nothing

End Sub

Public Sub Genera()
    Dim oComando As clsComando
    
    Screen.MousePointer = vbHourglass
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_TiempoSalon", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    fInicio = Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm")
    fFinal = Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm")
    oComando.CreateParameter "@flagTurnoOFecha", adBoolean, adParamInput, 1, chkTurno.Value
    oComando.CreateParameter "@tLocal", adVarChar, adParamInput, 20, cboLocal.BoundText
    oComando.CreateParameter "@tSalon", adVarChar, adParamInput, 20, cboSalon.BoundText
    oComando.CreateParameter "@tTurno", adVarChar, adParamInput, 20, sTurno
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set rsReporte = oComando.GetSP()
    Detallado.DiscardSavedData
    Detallado.Database.SetDataSource rsReporte
    Detallado.Text12.SetText "Tiempos en salon Detallado"
    Detallado.Text15.SetText sRazonSocial
    Detallado.Text9.SetText localConectado
    Detallado.ReportTitle = IIf(chkTurno.Value = 1, "Todos los Turnos", "Turno : " & sTurno) & sTitulo
    frmEmite.CRViewer.ReportSource = Detallado
    Screen.MousePointer = vbDefault
End Sub
