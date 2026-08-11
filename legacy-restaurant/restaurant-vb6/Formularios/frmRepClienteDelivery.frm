VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepClienteDelivery 
   Caption         =   "Emite de Clientes Frecuentes"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5835
   Icon            =   "frmRepClienteDelivery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
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
      Left            =   0
      Picture         =   "frmRepClienteDelivery.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   60
      Top             =   1605
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkDir 
      Caption         =   "Todos"
      Height          =   255
      Left            =   5010
      TabIndex        =   16
      Top             =   1890
      Value           =   1  'Checked
      Width           =   780
   End
   Begin VB.CheckBox chkZona 
      Caption         =   "Todos"
      Height          =   195
      Left            =   5010
      TabIndex        =   15
      Top             =   1530
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.TextBox txtDir 
      Height          =   345
      Left            =   975
      TabIndex        =   12
      ToolTipText     =   "Ingrese parte o toda la dirección"
      Top             =   1845
      Width           =   3915
   End
   Begin MSDataListLib.DataCombo cboZona 
      Height          =   315
      Left            =   975
      TabIndex        =   11
      ToolTipText     =   "Seleccione la zona de residencia del cliente"
      Top             =   1470
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "DataCombo1"
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
      Left            =   855
      Picture         =   "frmRepClienteDelivery.frx":082E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2280
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
      Left            =   2310
      Picture         =   "frmRepClienteDelivery.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2280
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
      Left            =   3720
      Picture         =   "frmRepClienteDelivery.frx":1292
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   " Activos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   2625
      Begin VB.OptionButton optActivo1 
         Caption         =   "Todos"
         Height          =   240
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Value           =   -1  'True
         Width           =   1545
      End
      Begin VB.OptionButton optActivo2 
         Caption         =   "Solo los Activos"
         Height          =   240
         Left            =   180
         TabIndex        =   6
         Top             =   615
         Width           =   1995
      End
      Begin VB.OptionButton optActivo3 
         Caption         =   "Solo los No Activos"
         Height          =   285
         Left            =   180
         TabIndex        =   5
         Top             =   960
         Width           =   2220
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Canjes por Puntos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   2685
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.OptionButton optCanje3 
         Caption         =   "Sin Canjes por Puntos"
         Height          =   285
         Left            =   180
         TabIndex        =   3
         Top             =   960
         Width           =   2475
      End
      Begin VB.OptionButton optCanje2 
         Caption         =   "Con Canjes por Puntos"
         Height          =   240
         Left            =   180
         TabIndex        =   2
         Top             =   660
         Width           =   1995
      End
      Begin VB.OptionButton optCanje1 
         Caption         =   "Todos"
         Height          =   240
         Left            =   180
         TabIndex        =   1
         Top             =   330
         Value           =   -1  'True
         Width           =   1545
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Dirección :"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Zona :"
      Height          =   195
      Left            =   420
      TabIndex        =   13
      Top             =   1530
      Width           =   465
   End
End
Attribute VB_Name = "frmRepClienteDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsZona      As Recordset
Dim RsReporte   As Recordset
Dim Reporte     As New dsrClienteDelivery

Dim xIsql As String

Sub LlenaCombos()
    With cboZona
         Isql = "Select Codigo, Descripcion from vZona where lActivo='1' order by Descripcion"
         Set RsZona = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsZona
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    If Index = 3 Then
        Unload Me
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Genera
    If RsReporte.RecordCount = 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "No hay datos que mostrar", vbInformation, sMensaje
        Exit Sub
    End If
    
    Reporte.DiscardSavedData
    Reporte.Database.SetDataSource RsReporte
    Reporte.ReportTitle = Me.Caption
    Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    Reporte.PaperSize = crPaperA4
    Reporte.PaperOrientation = crLandscape
    Reporte.Text1.SetText sRazonSocial
    Reporte.Text18.SetText localConectado
    Select Case Index
        Case Is = 0 ' Preview
             
             frmEmite.CRViewer.ReportSource = Reporte
             frmEmite.CRViewer.DisplayGroupTree = False
             frmEmite.CRViewer.ViewReport
             frmEmite.Show vbModal
             
        Case Is = 1 ' Imprimir
             Screen.MousePointer = vbDefault
             Reporte.PrintOut
        
        Case Is = 2 ' Exportar
               Screen.MousePointer = vbDefault
                  Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Reporte.PaperOrientation = crLandscape
                  
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
    Screen.MousePointer = vbDefault
    Exit Sub
Exportar:
    MsgBox err.Number & " " & err.Description & " " & err.Source, vbCritical
End Sub

Private Sub chkDir_Click()
    txtDir = ""
    If chkDir.Value = 1 Then
        txtDir.Enabled = False
    Else
        txtDir.Enabled = True
    End If
End Sub

Private Sub chkZona_Click()
    cboZona.Text = ""
    If chkZona.Value = 1 Then
        cboZona.Enabled = False
    Else
        cboZona.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Centrar Me
    xIsql = "0"
    LlenaCombos
    cboZona.Enabled = 0
    cboZona.Text = ""
    txtDir.Enabled = 0
    
End Sub

Public Sub Genera()
    'Consistencia q se ingrese la Zona
   If chkZona.Value = 0 Then
      If cboZona.Text = "" Then
         MsgBox "Debe escoger la zona", vbCritical, sMensaje
         Exit Sub
      End If
   End If
    'Consistencia q se ingrese la Dirección
    If chkDir.Value = 0 Then
        If txtDir = "" Then
            MsgBox "Debe ingresar una dirección", vbCritical, sMensaje
            Exit Sub
        End If
    End If
    
    xIsql = "0"
    Isql = ""
    
    'Flag Activo
    If Not optActivo1.Value Then
        If optActivo2.Value = True Then
           xIsql = "lActivo =1"
        End If
        If optActivo3.Value = True Then
           xIsql = "lActivo =0"
        End If
    End If
    
    'Flag Canje
    If Not optCanje1.Value Then
        If optCanje2.Value = True Then
           If xIsql = "0" Then
              xIsql = "lPuntos=1"
           Else
              xIsql = xIsql & " and lPuntos=1"
           End If
        End If
        If optCanje3.Value = True Then
           If xIsql = "0" Then
              xIsql = "lPuntos =0"
           Else
              xIsql = xIsql & " and lPuntos=0"
           End If
        End If
    End If
    'Zona
    If chkZona.Value = False Then
        If xIsql = "0" Then
            xIsql = " tZona='" & cboZona.BoundText & "'"
        Else
            xIsql = xIsql & " AND tZona='" & cboZona.BoundText & "'"
        End If
    End If
    'Dirección
    If chkDir.Value = False Then
        If xIsql = "0" Then
            xIsql = " tDireccion LIKE '%" & txtDir & "%'"
        Else
            xIsql = xIsql & " AND tDireccion LIKE '%" & txtDir & "%'"
        End If
    End If
    'consulta final q adiciona o nó el filtro
    If xIsql = "0" Then
        Isql = "SELECT tNumeroTarjeta as tTarjeta, tTelefono, LTrim(tApellido) + ' ' + LTrim(tNombre) as tApellido, tDireccion, tReferencia, nDescuento,(CASE WHEN dbo.vDelivery.lPuntos=1 THEN 'SI' ELSE 'NO' END ) AS Puntos, fNacimiento, tEMail, nAcumulado, nUtilizado, nDisponible FROM dbo.vDelivery "
    Else
        Isql = "SELECT tNumeroTarjeta as tTarjeta, tTelefono, LTrim(tApellido) + ' ' + LTrim(tNombre) as tApellido, tDireccion, tReferencia, nDescuento,(CASE WHEN dbo.vDelivery.lPuntos=1 THEN 'SI' ELSE 'NO' END ) AS Puntos, fNacimiento, tEMail, nAcumulado, nUtilizado, nDisponible FROM dbo.vDelivery " & _
               "WHERE " & xIsql
    End If
    Set RsReporte = Lib.OpenRecordset(Isql, Cn)

End Sub


