VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRepReservas 
   Caption         =   "Reservas"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5370
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
      Index           =   2
      Left            =   3840
      Picture         =   "frmRepReservas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2460
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
      Left            =   2400
      Picture         =   "frmRepReservas.frx":00F2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2460
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   30
      TabIndex        =   0
      Top             =   -15
      Width           =   5295
      Begin VB.CheckBox chkEstado 
         Caption         =   "Anulado"
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   13
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CheckBox chkEstado 
         Caption         =   "Atendido"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   12
         Top             =   1800
         Width           =   975
      End
      Begin VB.CheckBox chkEstado 
         Caption         =   "Generado"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   11
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.ComboBox cboCriterio 
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
         Left            =   1650
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   375
         Left            =   1650
         TabIndex        =   1
         Top             =   720
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
         Format          =   154796033
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   375
         Left            =   1650
         TabIndex        =   2
         Top             =   240
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
         Format          =   154796033
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   3600
         TabIndex        =   9
         Top             =   240
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
         Format          =   154796035
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   3600
         TabIndex        =   10
         Top             =   720
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
         Format          =   154796035
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin VB.Label Label4 
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
         Left            =   840
         TabIndex        =   14
         Top             =   1800
         Width           =   720
      End
      Begin VB.Label Label1 
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
         Left            =   300
         TabIndex        =   8
         Top             =   1320
         Width           =   1290
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
         Left            =   360
         TabIndex        =   4
         Top             =   330
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
         Left            =   465
         TabIndex        =   3
         Top             =   810
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmRepReservas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsCabecera As Recordset
Dim rsReporte As Recordset
Dim CriterioF As String
Dim CriterioS As String
Dim Reporte As New dsrReserva
Dim nColumna As Integer

Sub LlenaOrden()
    Dim i As Integer
    With cboCriterio
'         For i = 0 To grdGrilla.Columns.Count - 1
'             If grdGrilla.Columns(i).ValueItems.Presentation <> 4 Then
                .AddItem "Reserva"
                .ItemData(.NewIndex) = 0
                .AddItem "Nombres"
                .ItemData(.NewIndex) = 1
                .AddItem "Telefono"
                .ItemData(.NewIndex) = 2
                .AddItem "Fecha"
                .ItemData(.NewIndex) = 3
                .AddItem "Pax"
                .ItemData(.NewIndex) = 4
                .AddItem "Estado"
                .ItemData(.NewIndex) = 5
'             End If
'         Next i
    End With
    cboCriterio.SelText = "Reserva"
    CriterioS = "tReserva"
End Sub

Private Sub cboCriterio_Click()
    Select Case cboCriterio.ListIndex
        Case Is = 0
            CriterioS = "tReserva"
        Case Is = 1
            CriterioS = "tNombre"
        Case Is = 2
            CriterioS = "tTelefono"
        Case Is = 3
            CriterioS = "fFecha"
        Case Is = 4
            CriterioS = "nPax"
        Case Is = 5
            CriterioS = "tEstadoReserva"
    End Select
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
Select Case Index
    Case Is = 0  'Emite
        If dtpFecIni.value > dtpFecFin.value Then
            MsgBox "La Fecha Inicial no puede ser mayor que la fecha final", vbExclamation, "Mensaje del Sistema": Exit Sub
        End If
        If dtpFecIni.value = dtpFecFin.value Then
            If dtpHorIni.value > dtpHorFin.value Then
                MsgBox "La Fecha Inicial no puede ser mayor que la fecha final", vbExclamation, "Mensaje del Sistema": Exit Sub
            End If
        End If
        ObtenerReservas
        If RsCabecera.RecordCount = 0 Then MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema": Exit Sub
            Screen.MousePointer = vbHourglass
            Set rsReporte = RsCabecera.Clone
            rsReporte.Filter = CriterioF
            rsReporte.Sort = CriterioS & " ASC"
                            
            Reporte.DiscardSavedData
            Reporte.Database.SetDataSource rsReporte
            Reporte.Text4.SetText localConectado
            Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
            
            frmEmite.CRViewer.ReportSource = Reporte
            frmEmite.CRViewer.DisplayGroupTree = False
            frmEmite.CRViewer.ViewReport
            frmEmite.Show vbModal
    Case Is = 2  'Salida
        Unload Me
    End Select
End Sub

Private Sub ObtenerReservas()
    Screen.MousePointer = vbHourglass
    
    Centrar Me
    nColumna = 0
    
    ' Ingrese la Vista
'    Isql = "SELECT TRESERVA.tReserva, TRESERVA.fFecha+TRESERVA.fHora AS 'fFecha', TRESERVA.fHora, TRESERVA.tApellido, [tApellido]+' '+[tNombre] AS Cliente, TRESERVA.tNombre, TRESERVA.tTelefono, TRESERVA.nPax, TRESERVA.tEstadoReserva, TRESERVA.tObservacion, TRESERVA.fRegistro, vEstadoReserva.Descripcion AS EstadoReserva " & _
'           "FROM TRESERVA LEFT JOIN vEstadoReserva ON TRESERVA.tEstadoReserva = vEstadoReserva.Codigo Where (TRESERVA.fFecha + TRESERVA.fHora Between '" & Format(dtpFecIni.Value, "yyyy/mm/dd") + dtpHorIni.Value & "' AND '" & Format(dtpFecFin.Value, "yyyy/mm/dd") + dtpHorFin.Value & "') "
    Isql = "SELECT TRESERVA.tReserva, TRESERVA.fFecha+TRESERVA.fHora AS 'fFecha', TRESERVA.fHora, TRESERVA.tApellido, [tApellido]+' '+[tNombre] AS Cliente, TRESERVA.tNombre, TRESERVA.tTelefono, TRESERVA.nPax, TRESERVA.tEstadoReserva, TRESERVA.tObservacion, TRESERVA.fRegistro, UPPER(vEstadoReserva.Descripcion) AS EstadoReserva " & _
           "FROM TRESERVA LEFT JOIN vEstadoReserva ON TRESERVA.tEstadoReserva = vEstadoReserva.Codigo Where (TRESERVA.fFecha + TRESERVA.fHora Between '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm") & "' AND '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm") & "') "
    If chkEstado(0).value = 1 Or chkEstado(1).value = 1 Or chkEstado(2).value = 1 Then
        Dim IsqlEstado As String
        IsqlEstado = ""
        If chkEstado(0).value = 1 Then
            IsqlEstado = "'01',"
        End If
        If chkEstado(1).value = 1 Then
            IsqlEstado = IsqlEstado + "'02',"
        End If
        If chkEstado(2).value = 1 Then
            IsqlEstado = IsqlEstado + " '03',"
        End If
        Isql = Isql + " And TRESERVA.tEstadoReserva IN(" & Left(IsqlEstado, (Len(IsqlEstado) - 1)) & ")"
    End If
    
    Set RsCabecera = Lib.OpenRecordset(Isql, Cn)
    
    'RsCabecera.Filter = "tEstadoReserva ='01'"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Centrar Me
    LlenaOrden
    dtpFecIni.value = FechaServidor()
    dtpFecFin.value = FechaServidor()
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
End Sub
