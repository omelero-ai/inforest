VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRepClienteCtaCte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Emite de Clientes Cuentas Corrientes"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "frmRepClienteCtaCte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
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
      Left            =   2940
      Picture         =   "frmRepClienteCtaCte.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1905
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   " Empresa "
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
      TabIndex        =   11
      Top             =   0
      Width           =   3135
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   2580
         Top             =   225
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.OptionButton optEmpresa1 
         Caption         =   "Todos"
         Height          =   240
         Left            =   180
         TabIndex        =   14
         Top             =   330
         Value           =   -1  'True
         Width           =   1545
      End
      Begin VB.OptionButton optEmpresa2 
         Caption         =   "Solo las Empresas"
         Height          =   240
         Left            =   180
         TabIndex        =   13
         Top             =   660
         Width           =   1995
      End
      Begin VB.OptionButton optEmpresa3 
         Caption         =   "Solo las Personas Naturales"
         Height          =   285
         Left            =   180
         TabIndex        =   12
         Top             =   960
         Width           =   2490
      End
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
      TabIndex        =   7
      Top             =   0
      Width           =   2625
      Begin VB.OptionButton optActivo3 
         Caption         =   "Solo los No Activos"
         Height          =   285
         Left            =   180
         TabIndex        =   10
         Top             =   960
         Width           =   2220
      End
      Begin VB.OptionButton optActivo2 
         Caption         =   "Solo los Activos"
         Height          =   240
         Left            =   180
         TabIndex        =   9
         Top             =   615
         Width           =   1995
      End
      Begin VB.OptionButton optActivo1 
         Caption         =   "Todos"
         Height          =   240
         Left            =   180
         TabIndex        =   8
         Top             =   300
         Value           =   -1  'True
         Width           =   1545
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
      Left            =   4395
      Picture         =   "frmRepClienteCtaCte.frx":082E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1905
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
      Left            =   1485
      Picture         =   "frmRepClienteCtaCte.frx":0920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1905
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
      Left            =   30
      Picture         =   "frmRepClienteCtaCte.frx":0E52
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1905
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   45
      TabIndex        =   0
      Top             =   1410
      Width           =   5700
      Begin VB.CheckBox chkDir 
         Caption         =   "Todos"
         Height          =   255
         Left            =   4965
         TabIndex        =   2
         Top             =   105
         Value           =   1  'Checked
         Width           =   780
      End
      Begin VB.TextBox txtDir 
         Height          =   345
         Left            =   930
         TabIndex        =   1
         ToolTipText     =   "Ingrese parte o toda la dirección"
         Top             =   60
         Width           =   3915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Dirección :"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   135
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmRepClienteCtaCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsReporte   As Recordset
Dim Reporte     As New dsrClienteCtaCte

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
    Reporte.Text12.SetText localConectado
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
    If chkDir.value = 1 Then
        txtDir.Enabled = False
    Else
        txtDir.Enabled = True
    End If
End Sub

Private Sub Form_Load()
   Centrar Me
   xIsql = "0"
   txtDir.Enabled = 0
End Sub

Public Sub Genera()
    'Consistencia q se ingrese la Dirección
    If chkDir.value = 0 Then
        If txtDir = "" Then
            MsgBox "Debe ingresar una dirección", vbCritical, sMensaje
            Exit Sub
        End If
    End If
    
    xIsql = "0"
    Isql = ""
    
    If Not optActivo1.value Then
        If optActivo2.value = True Then
           xIsql = "lActivo =1"
        End If
        If optActivo3.value = True Then
           xIsql = "lActivo =0"
        End If
    End If
    '"Empresa"
    If Not optEmpresa1.value Then
        If optEmpresa2.value = True Then
           If xIsql = "0" Then
              xIsql = "lEmpresa=1"
           Else
              xIsql = xIsql & " and lEmpresa=1"
           End If
        End If
        If optEmpresa3.value = True Then
           If xIsql = "0" Then
              xIsql = "lEmpresa =0"
           Else
              xIsql = xIsql & " and lEmpresa=0"
           End If
        End If
    End If
    'Dirección
    If chkDir.value = False Then
        If xIsql = "0" Then
            xIsql = " tDireccion LIKE '%" & txtDir & "%'"
        Else
            xIsql = xIsql & " AND tDireccion LIKE '%" & txtDir & "%'"
        End If
    End If
    
    'consulta final q adiciona o nó el filtro
    If xIsql = "0" Then
        Isql = "SELECT Codigo, Identidad, Descripcion, tDireccion, tTelefono1, tEmail, nConsumo, nLinea FROM vCompania "
    Else
        Isql = "SELECT Codigo, Identidad, Descripcion, tDireccion, tTelefono1, tEmail, nConsumo, nLinea FROM vCompania " & _
               "WHERE " & xIsql
    End If
    
    Set RsReporte = Lib.OpenRecordset(Isql, Cn)

End Sub



