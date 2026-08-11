VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepCortesia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cortesías"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   Icon            =   "frmRepCortesia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Tipo Reporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   1680
      TabIndex        =   21
      Top             =   2280
      Width           =   4080
      Begin VB.OptionButton optForma 
         Caption         =   "Agrupacion Resumido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   28
         Top             =   1200
         Width           =   2775
      End
      Begin VB.OptionButton optForma 
         Caption         =   "Agrupacion (Sub Grupo)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   27
         Top             =   880
         Width           =   2775
      End
      Begin VB.OptionButton optForma 
         Caption         =   "Resumido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   600
         TabIndex        =   23
         Top             =   270
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.OptionButton optForma 
         Caption         =   "Detallado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   600
         TabIndex        =   22
         Top             =   585
         Width           =   1380
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
      Left            =   3840
      Picture         =   "frmRepCortesia.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5430
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
      Left            =   930
      Picture         =   "frmRepCortesia.frx":082E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5430
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
      Left            =   2385
      Picture         =   "frmRepCortesia.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5430
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
      Left            =   5295
      Picture         =   "frmRepCortesia.frx":1292
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5430
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
      Height          =   5385
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7665
      Begin VB.Frame Frame3 
         Caption         =   " Origen de Datos "
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
         Height          =   1245
         Left            =   3600
         TabIndex        =   24
         Top             =   3960
         Width           =   2130
         Begin VB.OptionButton optDatos 
            Caption         =   "Combos"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   585
            Width           =   1500
         End
         Begin VB.OptionButton optDatos 
            Caption         =   "Platos de Venta"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   270
            Value           =   -1  'True
            Width           =   1860
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " Valorización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   1680
         TabIndex        =   17
         Top             =   3960
         Width           =   1800
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Venta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   20
            Top             =   270
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Neto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   90
            TabIndex        =   19
            Top             =   585
            Width           =   1455
         End
         Begin VB.OptionButton optValor 
            Caption         =   "Precio Costo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   90
            TabIndex        =   18
            Top             =   900
            Width           =   1455
         End
      End
      Begin VB.CheckBox chkTipoPedido 
         Caption         =   "Todos los Pedidos"
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
         Left            =   5460
         TabIndex        =   3
         Top             =   840
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin VB.CheckBox chkCortesia 
         Caption         =   "Todas las Cortesías"
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
         Left            =   5460
         TabIndex        =   1
         Top             =   435
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   1245
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
         Format          =   56623107
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   3720
         TabIndex        =   7
         Top             =   1710
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
         Format          =   56623107
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   375
         Left            =   1710
         TabIndex        =   6
         Top             =   1710
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
         Format          =   56623105
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   375
         Left            =   1710
         TabIndex        =   4
         Top             =   1245
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
         Format          =   56623105
         CurrentDate     =   37539.2083333333
      End
      Begin MSDataListLib.DataCombo cboTipoPedido 
         Height          =   315
         Left            =   1710
         TabIndex        =   2
         Top             =   825
         Width           =   3645
         _ExtentX        =   6429
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
      Begin MSDataListLib.DataCombo cboCortesia 
         Height          =   315
         Left            =   1710
         TabIndex        =   0
         Top             =   420
         Width           =   3645
         _ExtentX        =   6429
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
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Pedido :"
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
         Left            =   195
         TabIndex        =   16
         Top             =   900
         Width           =   1425
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cortesia :"
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
         Left            =   795
         TabIndex        =   15
         Top             =   480
         Width           =   825
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
         Left            =   390
         TabIndex        =   14
         Top             =   1335
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
         Left            =   495
         TabIndex        =   13
         Top             =   1755
         Width           =   1125
      End
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   120
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepCortesia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNombre As String
Dim rsReporte As Recordset
Dim RsCortesia As Recordset
Dim RsTipoPedido As Recordset
Dim Reporte As New dsrCortesia
Dim ReporteD As New dsrCortesiaD
Dim ReporteA As New dsrCortesiaA
Dim ReporteAR As New dsrCortesiaAR
Dim sCriterio As String
Dim sTitulo As String

Dim fInicio As Date
Dim fFinal As Date

Sub LlenaCombos()
    With cboTipoPedido
         Isql = "Select * from vTipoPedido"
         Set RsTipoPedido = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoPedido
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With

    With cboCortesia
         Isql = "Select * from vCortesia"
         Set RsCortesia = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsCortesia
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
End Sub

Private Sub chkTipoPedido_Click()
   If chkTipoPedido.Value = 1 Then
      cboTipoPedido.Enabled = False
      cboTipoPedido.Text = ""
   Else
      cboTipoPedido.Enabled = True
   End If
End Sub

Private Sub chkCortesia_Click()
   If chkCortesia.Value = 1 Then
      cboCortesia.Enabled = False
      cboCortesia.Text = ""
   Else
      cboCortesia.Enabled = True
   End If
End Sub

Private Sub cmdOpcion_Click(index As Integer)
 
    If index = 2 Then
        Unload Me
        Exit Sub
    End If
   sCriterio = ""
   sTitulo = ""

   ' sCriterio = "MDOCUMENTO.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm") & "' and MDOCUMENTO.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm") & "'"
    sTitulo = " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.Value, "HH:nn") & " al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.Value, "HH:nn")
    
    If chkCortesia.Value = 0 Then
        If cboCortesia.Text = "" Then
            MsgBox "Debe escoger un Cortesia", vbCritical, sMensaje
            Exit Sub
        End If
        sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & " tCortesia ='" & cboCortesia.BoundText & "'"
    End If
        
    If chkTipoPedido.Value = 0 Then
        If cboTipoPedido.Text = "" Then
            MsgBox "Debe escoger un Tipo de Pedido", vbCritical, sMensaje
            Exit Sub
        End If
        sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & " tTipoPedido ='" & cboTipoPedido.BoundText & "'"
    End If
        
    Select Case index
        Case Is = 0 ' Preview
            Genera
            If rsReporte.EOF = True Then
                MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            frmEmite.CRViewer.DisplayGroupTree = False
            If optForma(0).Value Then
                Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                Reporte.PaperOrientation = crPortrait
            ElseIf optForma(1).Value Then
                ReporteD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                ReporteD.PaperOrientation = crPortrait
            ElseIf optForma(2).Value Then
                ReporteA.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                ReporteA.PaperOrientation = crPortrait
            Else
                ReporteAR.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                ReporteAR.PaperOrientation = crPortrait
            End If
            frmEmite.CRViewer.ViewReport
            frmEmite.Show vbModal
        
        Case Is = 1 ' Imprimir
            Genera
            Screen.MousePointer = vbDefault
            If rsReporte.EOF = True Then
                MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                Exit Sub
            End If
            If optForma(0).Value Then
                Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                Reporte.PaperOrientation = crPortrait
                Reporte.PrintOut
            ElseIf optForma(1).Value Then
                ReporteD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                ReporteD.PaperOrientation = crPortrait
                ReporteD.PrintOut
            ElseIf optForma(2).Value Then
                ReporteA.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                ReporteA.PaperOrientation = crPortrait
                ReporteA.PrintOut
            Else
                ReporteAR.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                ReporteAR.PaperOrientation = crPortrait
                ReporteAR.PrintOut
            End If
        
        Case Is = 2 ' Salir
            Unload Me
        
        Case Is = 3 ' Exportar
            Genera index
            If rsReporte.EOF = True Then
                MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                Exit Sub
            End If
            rsReporte.MoveFirst
            ExportaExcel rsReporte
            Screen.MousePointer = vbDefault
            
'            If rsReporte.EOF = True Then
'               MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
'               Exit Sub
'            End If
'
'            If optForma(0).Value Then
'                Reporte.ExportOptions.FormatType = 21
'                Reporte.ExportOptions.DestinationType = 1
'            ElseIf optForma(1).Value Then
'                ReporteD.ExportOptions.FormatType = 21
'                ReporteD.ExportOptions.DestinationType = 1
'            ElseIf optForma(2).Value Then
'                ReporteA.ExportOptions.FormatType = 21
'                ReporteA.ExportOptions.DestinationType = 1
'            Else
'                ReporteAR.ExportOptions.FormatType = 21
'                ReporteAR.ExportOptions.DestinationType = 1
'            End If
'
'            cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
'            cmdSave.ShowSave
'            If cmdSave.FileName = "" Then
'               Exit Sub
'            End If
'            If optForma(0).Value Then
'                Reporte.ExportOptions.DiskFileName = cmdSave.FileName
'                Reporte.Export False
'            ElseIf optForma(1).Value Then
'                ReporteD.ExportOptions.DiskFileName = cmdSave.FileName
'                ReporteD.Export False
'            ElseIf optForma(2).Value Then
'                ReporteA.ExportOptions.DiskFileName = cmdSave.FileName
'                ReporteA.Export False
'            Else
'                ReporteAR.ExportOptions.DiskFileName = cmdSave.FileName
'                ReporteAR.Export False
'            End If
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
    dtpFecIni.Value = Date
    dtpFecFin.Value = Date
    cboCortesia.Enabled = False
    cboTipoPedido.Enabled = False
    cboTipoPedido.Text = ""
    cboCortesia.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
   Set RsCortesia = Nothing
End Sub

Public Sub Genera(Optional index As Integer)
    Dim oComando As clsComando
    
    Dim reporteForma As String
    
    Dim TipoEmision As String
    
    If dtpFecIni.Value + dtpHorIni.Value > dtpFecFin.Value + dtpHorFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      Exit Sub
      dtpFecFin.SetFocus
    End If
    
    fInicio = Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm")
    fFinal = Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm")
    
    If optValor(0).Value = True Then
        reporteForma = "1" '  valor venta
        sTitulo = sTitulo & "Precio Venta"
    ElseIf optValor(1).Value = True Then
        reporteForma = "2"
        sTitulo = sTitulo & "Precio Neto"
    ElseIf optValor(2).Value = True Then
        reporteForma = "3"
        sTitulo = sTitulo & "Precio Costo"
    End If
    
    'TIPO EMISION
    If optForma(0).Value = True Then
        TipoEmision = "RESUMIDO"
    ElseIf optForma(1).Value = True Then
        TipoEmision = "DETALLADO"
    ElseIf optForma(2).Value = True Then
        TipoEmision = "AGRUPADO"
    ElseIf optForma(3).Value = True Then
        TipoEmision = "AGRUPADO"
    End If
            
            
            
            
    Screen.MousePointer = vbHourglass
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_Cortesia", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    oComando.CreateParameter "@tipo", adVarChar, adParamInput, 10, TipoEmision
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@Forma", adVarChar, adParamInput, 1, reporteForma
    oComando.CreateParameter "@tipoX", adVarChar, adParamInput, 10, IIf(Me.optDatos(0).Value = True, "VENTA", "COMBO")
    oComando.CreateParameter "@sCriterio", adVarChar, adParamInput, 400, sCriterio
        
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set rsReporte = oComando.GetSP()
    'rsReporte.Filter = sCriterio
     If index = 3 Then
        Exit Sub
     End If
    If optForma(0).Value Then

        Reporte.DiscardSavedData
        Reporte.Database.SetDataSource rsReporte
        Reporte.Text15.SetText sRazonSocial
        If optValor(0).Value = True Then
            Reporte.Text4.SetText "Venta"
        ElseIf optValor(1).Value = True Then
            Reporte.Text4.SetText "Neto"
        ElseIf optValor(2).Value = True Then
            Reporte.Text4.SetText "Costo"
        End If
        
        Reporte.Text10.SetText localConectado
        Reporte.ReportTitle = IIf(chkCortesia.Value = 1, "Todas las Cortesías ", "Cortesía : " & cboCortesia.Text) & sTitulo
        frmEmite.CRViewer.ReportSource = Reporte
        
    ElseIf optForma(1).Value = True Then

        ReporteD.DiscardSavedData
        ReporteD.Database.SetDataSource rsReporte
        ReporteD.Text15.SetText sRazonSocial
        ReporteD.Text1.SetText localConectado
        ReporteD.ReportTitle = IIf(chkCortesia.Value = 1, "Todas las Cortesías ", "Cortesía : " & cboCortesia.Text) & sTitulo
        frmEmite.CRViewer.ReportSource = ReporteD
    
    ElseIf optForma(2).Value = True Then
    
        ReporteA.DiscardSavedData
        ReporteA.Database.SetDataSource rsReporte
        ReporteA.Text15.SetText sRazonSocial
        ReporteA.Text1.SetText localConectado
        ReporteA.ReportTitle = IIf(chkCortesia.Value = 1, "Todas las Cortesías ", "Cortesía : " & cboCortesia.Text) & " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.Value, "HH:nn") & " al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.Value, "HH:nn")
        frmEmite.CRViewer.ReportSource = ReporteA
    
    Else
        ReporteAR.DiscardSavedData
        ReporteAR.Database.SetDataSource rsReporte
        ReporteAR.Text15.SetText sRazonSocial
        ReporteAR.Text1.SetText localConectado
        ReporteAR.ReportTitle = IIf(chkCortesia.Value = 1, "Todas las Cortesías ", "Cortesía : " & cboCortesia.Text) & " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.Value, "HH:nn") & " al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.Value, "HH:nn")
        frmEmite.CRViewer.ReportSource = ReporteAR
    End If
    
End Sub

Private Sub optforma_Click(index As Integer)
    If optForma(1).Value = True Then
       Frame3.Enabled = True
       optDatos(0).Enabled = True
       optDatos(1).Enabled = True
       Frame4.Enabled = True
       optValor(0).Enabled = True
       optValor(1).Enabled = True
       optValor(2).Enabled = True
       
       cboCortesia.Enabled = True
       cboTipoPedido.Enabled = True
       chkCortesia.Enabled = True
       chkTipoPedido.Enabled = True
       chkCortesia.Value = 1
       chkTipoPedido.Value = 1
       
    ElseIf optForma(0).Value = True Then
       Frame3.Enabled = False
       optDatos(0).Enabled = False
       optDatos(1).Enabled = False
       Frame4.Enabled = True
       optValor(0).Enabled = True
       optValor(1).Enabled = True
       optValor(2).Enabled = True
       
       cboCortesia.Enabled = True
       cboTipoPedido.Enabled = True
       chkCortesia.Enabled = True
       chkTipoPedido.Enabled = True
       chkCortesia.Value = 1
       chkTipoPedido.Value = 1
       
    ElseIf optForma(2).Value = True Then
       Frame4.Enabled = False
       optValor(0).Enabled = False
       optValor(1).Enabled = False
       optValor(2).Enabled = False
       Frame3.Enabled = True
       optDatos(0).Enabled = True
       optDatos(1).Enabled = True
       
       cboCortesia.Enabled = False
       cboTipoPedido.Enabled = False
       chkCortesia.Enabled = False
       chkTipoPedido.Enabled = False
       chkCortesia.Value = 1
       chkTipoPedido.Value = 1
       cboCortesia.Text = ""
       cboTipoPedido.Text = ""
       
    Else
       Frame4.Enabled = False
       optValor(0).Enabled = False
       optValor(1).Enabled = False
       optValor(2).Enabled = False
       Frame3.Enabled = False
       optDatos(0).Enabled = False
       optDatos(1).Enabled = False
    
       cboCortesia.Enabled = False
       cboTipoPedido.Enabled = False
       chkCortesia.Enabled = False
       chkTipoPedido.Enabled = False
       chkCortesia.Value = 1
       chkTipoPedido.Value = 1
       cboCortesia.Text = ""
       cboTipoPedido.Text = ""
    End If
    
End Sub
