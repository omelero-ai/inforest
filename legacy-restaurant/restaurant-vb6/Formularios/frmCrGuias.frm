VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCrGuias 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   Icon            =   "frmCrGuias.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   4920
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   810
      Left            =   150
      TabIndex        =   25
      Top             =   2550
      Width           =   4585
      Begin VB.OptionButton optChofer 
         Caption         =   "Todos"
         Height          =   270
         Index           =   1
         Left            =   3465
         TabIndex        =   31
         Top             =   45
         Value           =   -1  'True
         Width           =   930
      End
      Begin VB.OptionButton optChofer 
         Caption         =   "Seleccionar"
         Height          =   270
         Index           =   0
         Left            =   1305
         TabIndex        =   30
         Top             =   60
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo dcboChofer 
         Height          =   315
         Left            =   870
         TabIndex        =   26
         Top             =   360
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   15788763
         Text            =   ""
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Chofer :"
         Height          =   195
         Left            =   30
         TabIndex        =   27
         Top             =   45
         Width           =   555
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   135
      TabIndex        =   22
      Top             =   1755
      Width           =   4585
      Begin VB.OptionButton optTransportista 
         Caption         =   "Todos"
         Height          =   270
         Index           =   1
         Left            =   3465
         TabIndex        =   29
         Top             =   0
         Value           =   -1  'True
         Width           =   930
      End
      Begin VB.OptionButton optTransportista 
         Caption         =   "Seleccionar"
         Height          =   270
         Index           =   0
         Left            =   1305
         TabIndex        =   28
         Top             =   15
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo dcboTransportista 
         Height          =   315
         Left            =   870
         TabIndex        =   23
         Top             =   315
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   15788763
         Text            =   ""
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Transportista :"
         Height          =   195
         Left            =   30
         TabIndex        =   24
         Top             =   45
         Width           =   1005
      End
   End
   Begin VB.OptionButton optDestinatario 
      Caption         =   "Seleccionar"
      Height          =   270
      Index           =   0
      Left            =   1440
      TabIndex        =   21
      Top             =   990
      Width           =   1215
   End
   Begin VB.OptionButton optDestinatario 
      Caption         =   "Todos"
      Height          =   270
      Index           =   1
      Left            =   3600
      TabIndex        =   20
      Top             =   975
      Value           =   -1  'True
      Width           =   930
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   135
      TabIndex        =   17
      Top             =   990
      Width           =   4585
      Begin MSDataListLib.DataCombo dcboDestinatario 
         Height          =   315
         Left            =   870
         TabIndex        =   18
         Top             =   270
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   15788763
         Text            =   ""
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Destinatario :"
         Height          =   195
         Left            =   30
         TabIndex        =   19
         Top             =   45
         Width           =   930
      End
   End
   Begin VB.Frame Frame10 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   135
      TabIndex        =   11
      Top             =   3555
      Width           =   4485
      Begin MSComCtl2.DTPicker DtpInicio 
         Height          =   315
         Left            =   870
         TabIndex        =   3
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   15788763
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51970051
         CurrentDate     =   37783
      End
      Begin MSComCtl2.DTPicker DtpFin 
         Height          =   315
         Left            =   3240
         TabIndex        =   4
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   15788763
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   16777215
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51970051
         CurrentDate     =   37783
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio :"
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   60
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fin :"
         Height          =   195
         Left            =   2760
         TabIndex        =   12
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.Frame Frame11 
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   180
      TabIndex        =   9
      Top             =   4215
      Width           =   4485
      Begin VB.OptionButton optSalida 
         Caption         =   "Impresora"
         Height          =   225
         Index           =   1
         Left            =   2160
         TabIndex        =   6
         Top             =   75
         Width           =   1035
      End
      Begin VB.OptionButton optSalida 
         Caption         =   "Vista Previa"
         Height          =   225
         Index           =   0
         Left            =   840
         TabIndex        =   5
         Top             =   60
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.OptionButton optSalida 
         Caption         =   "Archivo XLS"
         Height          =   225
         Index           =   2
         Left            =   3290
         TabIndex        =   7
         Top             =   75
         Width           =   1185
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Salida :"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   60
         Width           =   525
      End
   End
   Begin VB.Frame Frame4 
      Height          =   120
      Left            =   -105
      TabIndex        =   8
      Top             =   4785
      Width           =   5190
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   -60
      TabIndex        =   1
      Top             =   340
      Width           =   5055
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   945
      Picture         =   "frmCrGuias.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5025
      UseMaskColor    =   -1  'True
      Width           =   1150
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2700
      Picture         =   "frmCrGuias.frx":068C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5025
      UseMaskColor    =   -1  'True
      Width           =   1150
   End
   Begin MSComDlg.CommonDialog dlgExportar 
      Left            =   4275
      Top             =   5025
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   135
      X2              =   4815
      Y1              =   3375
      Y2              =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   135
      X2              =   4815
      Y1              =   2535
      Y2              =   2535
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4800
      Y1              =   1725
      Y2              =   1725
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   4800
      Y1              =   4095
      Y2              =   4095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3855
      Left            =   90
      TabIndex        =   16
      Top             =   900
      Width           =   4725
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Agrupando por :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   15
      Top             =   555
      Width           =   1545
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Guías"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   105
      TabIndex        =   14
      Top             =   0
      Width           =   810
   End
End
Attribute VB_Name = "frmCrGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsDestinatario      As New ADODB.Recordset
Dim RsTransportista     As New ADODB.Recordset
Dim RsChofer            As New ADODB.Recordset
Dim RsReporte           As New ADODB.Recordset
Dim Reporte             As New dsrGuias

Dim xFiltro             As String
Dim NomReporte          As String

Sub Libera()
    Set RsDestinatario = Nothing
    Set RsTransportista = Nothing
    Set RsChofer = Nothing
    Set Reporte = Nothing
End Sub

Sub LlenaCombo()
    'Destinatario
    Isql = "SELECT Codigo, Descripcion FROM dbo.vCliente WHERE LACTIVO='1'ORDER BY Codigo"
    Llena_Combo Isql, dcboDestinatario, RsDestinatario, "Codigo", "Descripcion"
    dcboDestinatario.BoundText = IIf(RsDestinatario.RecordCount > 0, RsDestinatario!Codigo, "")
    'Transportista
    'porq son los mismo datos que el destinatario, por eos no necesito una nueva ISQL
    'Isql = "SELECT Codigo, Descripcion  FROM vCliente WHERE LACTIVO=1 ORDER BY Codigo"
    Llena_Combo Isql, dcboTransportista, RsTransportista, "Codigo", "Descripcion"
    dcboTransportista.BoundText = IIf(RsTransportista.RecordCount > 0, RsTransportista!Codigo, "")
    'Chofer
    Isql = "SELECT * FROM vChofer WHERE LACTIVO=1 ORDER BY Codigo"
    Llena_Combo Isql, dcboChofer, RsChofer, "Codigo", "Descripcion"
    dcboChofer.BoundText = IIf(RsChofer.RecordCount > 0, RsChofer!Codigo, "")
    
End Sub

Sub Mostrar()
    Dim fechaTemp As Date
   
    xFiltro = ""
    xFiltro = "fFecha >= '" & Format(DtpInicio.Value, "YYYY-MM-DD") & " 00:00' AND fFecha <= '" & Format(DtpFin.Value, "YYYY-mm-dd") & " 23:59' "
    If optDestinatario(0).Value Then
        xFiltro = xFiltro & "AND tCodDestinatario='" & dcboDestinatario.BoundText & "'"
    End If
    If optTransportista(0).Value Then
        xFiltro = xFiltro & "AND tCodTransportista='" & dcboTransportista.BoundText & "'"
    End If
    If optChofer(0).Value Then
        xFiltro = xFiltro & "AND tCodChofer='" & dcboChofer.BoundText & "'"
    End If
    
    'Emite el reporte de guías
    'NomReporte = "CrGuias.rpt"
    Isql = "SELECT tGuia, tDocumento, nPeso, nTara, nTotal, Destinatario, Transportista, Vehiculo, Chofer, Estado, Motivo FROM vGuia WHERE " & xFiltro
    Set RsReporte = Lib.OpenRecordset(Isql, Cn)

End Sub

Private Sub cmdAceptar_Click()
    If Len(Trim(dcboDestinatario.BoundText)) = 0 Then MsgBox "Seleccione el Destinatario", vbExclamation, sMensaje: dcboDestinatario.SetFocus: Exit Sub
    If Len(Trim(dcboTransportista.BoundText)) = 0 Then MsgBox "Seleccione el transportista", vbExclamation, sMensaje: dcboTransportista.SetFocus: Exit Sub
    If Len(Trim(dcboChofer.BoundText)) = 0 Then MsgBox "Seleccione el Chofer", vbExclamation, sMensaje: dcboChofer.SetFocus: Exit Sub
    
    Screen.MousePointer = vbHourglass
    Call Mostrar
    
    If RsReporte = 0 Then MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema": Exit Sub
    Screen.MousePointer = vbHourglass

    Set RsReporte = RsReporte
    
    Reporte.DiscardSavedData
    Reporte.ReportTitle = Me.Caption
'    Reporte.Text3.SetText "Numero Brevete"
'    Reporte.Text7.SetText "Estado"
    Reporte.Database.SetDataSource RsReporte
    Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    frmEmite.CRViewer.ReportSource = Reporte
    frmEmite.CRViewer.DisplayGroupTree = False
    frmEmite.CRViewer.ViewReport
    frmEmite.Show vbModal

    
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub DtpFin_Change()
    If DateValue(DtpFin.Value) < DateValue(DtpInicio.Value) Then
        MsgBox "La fecha de final no puede ser menor a la fecha de inicio", vbExclamation, Mensaje
        DtpFin.Value = DateValue(DtpInicio.Value) + 1
    End If
End Sub

Private Sub DtpInicio_Change()
    If DateValue(DtpFin.Value) < DateValue(DtpInicio.Value) Then
        MsgBox "La fecha de inicio no puede ser menor a la fecha final", vbExclamation, Mensaje
        DtpInicio.Value = DateValue(DtpFin.Value) - 1
    End If
End Sub

Private Sub Form_Load()
    Centrar Me
    Call LlenaCombo
    
    
    DtpInicio = Date
    DtpFin = Date + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Libera
'    Set frmCrMTCMovViajeros = Nothing
End Sub

