VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepDescuento 
   Caption         =   "Descuentos"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8940
   Icon            =   "frmRepDescuento.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Exportar Grillado"
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
      Left            =   4590
      Picture         =   "frmRepDescuento.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4905
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
      Left            =   1680
      Picture         =   "frmRepDescuento.frx":082E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4905
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
      Left            =   3135
      Picture         =   "frmRepDescuento.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4905
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
      Left            =   6045
      Picture         =   "frmRepDescuento.frx":1292
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4905
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
      Height          =   4875
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   8835
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
         Height          =   1335
         Left            =   2040
         TabIndex        =   20
         Top             =   2160
         Width           =   3495
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
            Index           =   3
            Left            =   240
            TabIndex        =   23
            Top             =   840
            Width           =   3135
         End
         Begin VB.OptionButton optForma 
            Caption         =   "Pedido Detallado"
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
            Index           =   1
            Left            =   240
            TabIndex        =   22
            Top             =   600
            Width           =   2415
         End
         Begin VB.OptionButton optForma 
            Caption         =   "Pedido Resumido"
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
            Index           =   0
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Value           =   -1  'True
            Width           =   2415
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " Valores "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2040
         TabIndex        =   19
         Top             =   3600
         Width           =   3495
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
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   14
            Top             =   480
            Width           =   1455
         End
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
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.CheckBox chkTipoPedido 
         Caption         =   "Todos los Tipos de Pedidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5685
         TabIndex        =   8
         Top             =   765
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkDescuento 
         Caption         =   "Todos los Tipos de Descuentos"
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
         Left            =   5685
         TabIndex        =   6
         Top             =   420
         Value           =   1  'Checked
         Width           =   3075
      End
      Begin MSDataListLib.DataCombo cboDescuento 
         Height          =   315
         Left            =   2025
         TabIndex        =   5
         Top             =   450
         Width           =   3585
         _ExtentX        =   6324
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
      Begin MSDataListLib.DataCombo cboTipoPedido 
         Height          =   315
         Left            =   2025
         TabIndex        =   7
         Top             =   810
         Width           =   3585
         _ExtentX        =   6324
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
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   375
         Left            =   2025
         TabIndex        =   9
         Top             =   1215
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
         Format          =   56033281
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   4110
         TabIndex        =   10
         Top             =   1215
         Width           =   1500
         _ExtentX        =   2646
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
         Format          =   56033283
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   375
         Left            =   2025
         TabIndex        =   11
         Top             =   1665
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
         Format          =   56033281
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   4110
         TabIndex        =   12
         Top             =   1665
         Width           =   1500
         _ExtentX        =   2646
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
         Format          =   56033283
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
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
         Left            =   510
         TabIndex        =   4
         Top             =   855
         Width           =   1425
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Descuento :"
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
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   480
         Width           =   1755
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
         Left            =   705
         TabIndex        =   2
         Top             =   1290
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
         Left            =   810
         TabIndex        =   1
         Top             =   1710
         Width           =   1125
      End
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   900
      Top             =   4950
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepDescuento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNombre As String
Dim rsReporte As Recordset
Dim RsMotivoDescuento As Recordset
Dim RsTipoPedido As Recordset
Dim Reporte As New dsrDescuento
Dim ReporteD As New dsrDescuentoD
Dim sTemp   As String
Dim sTempD   As String
Dim sCriterio As String
Dim scriterio1 As String
Dim sCriterio2 As String
Dim sTitulo As String
Dim RsTemporal As Recordset
Dim coddoc As String
Dim desc As Double
Dim sPrecio As String
Dim sTexto As String

Dim fInicio As Date
Dim fFinal As Date
Dim tipoDescuento As String
Dim tipopedido As String

'CESAR 0020-2013
Dim sTipo As String
Dim sValor As String
Dim ReporteA As New dsrDescuentoA
Dim sComp As String

Sub LlenaCombos()
    With cboTipoPedido
         Isql = "Select * from vTipoPedido"
         Set RsTipoPedido = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoPedido
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With

    With cboDescuento
         Isql = "Select * from vMotivoDescuento"
         Set RsMotivoDescuento = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMotivoDescuento
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
End Sub

Private Sub chkDescuento_Click()
   If chkDescuento.Value = 1 Then
      cboDescuento.Enabled = False
      cboDescuento.Text = ""
   Else
      cboDescuento.Enabled = True
   End If

End Sub

Private Sub chkTipoPedido_Click()
   If chkTipoPedido.Value = 1 Then
      cboTipoPedido.Enabled = False
      cboTipoPedido.Text = ""
   Else
      cboTipoPedido.Enabled = True
   End If
End Sub


Private Sub cmdOpcion_Click(Index As Integer)
   
    If Index = 2 Then
        Unload Me
        Exit Sub
    End If
    sCriterio = ""
    sCriterio2 = ""
    scriterio1 = ""
    tipoDescuento = ""
    tipopedido = ""
    
    sCriterio = "Mpedido.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm") & "' and Mpedido.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm") & "'"
    sCriterio2 = sCriterio
    scriterio1 = "Mdocumento.fRegistro >= '" & Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm") & "' and Mdocumento.fRegistro <= '" & Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm") & "'"
    sTitulo = " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.Value, "HH:nn") & " al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.Value, "HH:nn")
    
    If chkDescuento.Value = 0 Then
        If cboDescuento.Text = "" Then
            MsgBox "Debe escoger un Descuento", vbCritical, sMensaje
            Exit Sub
        End If
        sCriterio = sCriterio & " and vmotivodescuento.codigo ='" & cboDescuento.BoundText & "'"
        tipoDescuento = cboDescuento.BoundText
    End If
        
    If chkTipoPedido.Value = 0 Then
        If cboTipoPedido.Text = "" Then
            MsgBox "Debe escoger un Tipo de Pedido", vbCritical, sMensaje
            Exit Sub
        End If
        sCriterio = sCriterio & " and MPEDIDO.tTipoPedido ='" & cboTipoPedido.BoundText & "'"
        tipopedido = cboTipoPedido.BoundText
    End If
        
    Select Case Index
        Case Is = 0 ' Preview
            Genera
            If rsReporte.EOF = True Then
                MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            If optForma(0).Value Then 'Pedido Resumido
                frmEmite.CRViewer.DisplayGroupTree = True
                Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                Reporte.PaperOrientation = crPortrait
                
            ElseIf optForma(1).Value Then 'Pedido Detallado
                frmEmite.CRViewer.DisplayGroupTree = True
                ReporteD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                ReporteD.PaperOrientation = crPortrait
            
            Else 'Agrupado
                frmEmite.CRViewer.DisplayGroupTree = True
                ReporteA.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                ReporteA.PaperOrientation = crLandscape
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
            If optForma(0).Value Then 'Pedido Resumido
                Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                Reporte.PaperOrientation = crPortrait
                Reporte.PrintOut
                
            ElseIf optForma(1).Value Then 'Pedido Detallado
                ReporteD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                ReporteD.PaperOrientation = crPortrait
                ReporteD.PrintOut
                
            Else 'Agrupado
                ReporteA.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                ReporteA.PaperOrientation = crLandscape
                ReporteA.PrintOut
                
            End If
        
        Case Is = 2 ' Salir
            Unload Me
        
        Case Is = 3 ' Exportar
            If optForma(1).Value = False Then
                MsgBox "Proceso solo para tipo reporte detallado!!!", vbExclamation, "Mensaje del Sistema"
                Exit Sub
            End If
            Genera 1
            Screen.MousePointer = vbDefault
            If rsReporte.EOF = True Then
               MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
               Exit Sub
            End If
            
            ExportaExcel rsReporte
'            Genera
'            Screen.MousePointer = vbDefault
'            If rsReporte.EOF = True Then
'               MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
'               Exit Sub
'            End If
'
'            If optForma(0).Value Then 'Pedido Resumido
'                Reporte.ExportOptions.FormatType = 21
'                Reporte.ExportOptions.DestinationType = 1
'
'            ElseIf optForma(1).Value Then 'Pedido Detallado
'                ReporteD.ExportOptions.FormatType = 21
'                ReporteD.ExportOptions.DestinationType = 1
'
'            Else  'Agrupado
'                ReporteA.ExportOptions.FormatType = 21
'                ReporteA.ExportOptions.DestinationType = 1
'            End If
'
'            cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
'            cmdSave.ShowSave
'            If cmdSave.FileName = "" Then
'               Exit Sub
'            End If
'
'            If optForma(0).Value Then
'                Reporte.ExportOptions.DiskFileName = cmdSave.FileName
'                Reporte.Export False
'            ElseIf optForma(1).Value Then 'Pedido Detallado
'                ReporteD.ExportOptions.DiskFileName = cmdSave.FileName
'                ReporteD.Export False
'            Else  'Agrupado
'                ReporteA.ExportOptions.DiskFileName = cmdSave.FileName
'                ReporteA.Export False
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
    cboDescuento.Enabled = False
    cboTipoPedido.Enabled = False
    cboTipoPedido.Text = ""
    cboDescuento.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
   Set RsMotivoDescuento = Nothing
End Sub

Public Sub Genera(Optional ExportaExcel As Integer = 0)
    Screen.MousePointer = vbHourglass
    Dim oComando As clsComando
    Dim tipopedido As String
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_Descuento", Cn) Then
        Set oComando = Nothing
        Exit Sub
    End If
    If dtpFecIni.Value + dtpHorIni.Value > dtpFecFin.Value + dtpHorFin.Value Then
        MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
        Exit Sub
        dtpFecFin.SetFocus
    End If
    
    'CESAR 0020-2013
    If optValor(0).Value Then
       sValor = "01"
    ElseIf optValor(1).Value Then
       sValor = "02"
    Else
       sValor = "03"
    End If
    '---------------
    
    'CESAR 0020-2013  TIPO REPORTE
    If optForma(0).Value Then
       sTipo = "00"
    ElseIf optForma(1).Value Then
       sTipo = "01"
    Else
       sTipo = "02"
    End If
    '---------------
    
    fInicio = Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm")
    fFinal = Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm")
    If Me.cboTipoPedido.Text <> "" And Me.chkTipoPedido.Value = 0 Then
        tipopedido = cboTipoPedido.BoundText
        Else
        tipopedido = ""
    End If
        oComando.CreateParameter "@flag1", adVarChar, adParamInput, 2, sTipo
        oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
        oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
        oComando.CreateParameter "@mdesc", adVarChar, adParamInput, 10, tipoDescuento
        oComando.CreateParameter "@tipoPedido", adVarChar, adParamInput, 10, tipopedido
        'CESAR 0020-2013
        oComando.CreateParameter "@sValor", adVarChar, adParamInput, 50, sValor
        oComando.CreateParameter "@ExportaExcel", adBoolean, adParamInput, 1, ExportaExcel

   If Not oComando.GetParamOK Then
      Set oComando = Nothing
      Exit Sub
   End If
   Set rsReporte = oComando.GetSP()
   
   If ExportaExcel = 1 Then Exit Sub


    If sTipo = "00" Then 'Pedidos Resumido
          Reporte.DiscardSavedData
          Reporte.Database.SetDataSource rsReporte
          Reporte.Text5.SetText sRazonSocial
          Reporte.Text23.SetText localConectado
          Reporte.ReportTitle = sTitulo & sTexto
          frmEmite.CRViewer.ReportSource = Reporte
          
    ElseIf sTipo = "01" Then 'Pedidos Detallado
        ReporteD.DiscardSavedData
        ReporteD.Database.SetDataSource rsReporte
        ReporteD.Text5.SetText sRazonSocial
        ReporteD.Text24.SetText localConectado
        ReporteD.ReportTitle = sTitulo & sTexto
        frmEmite.CRViewer.ReportSource = ReporteD
        
    Else 'Agrupado
        ReporteA.DiscardSavedData
        ReporteA.Database.SetDataSource rsReporte
        ReporteA.Text13.SetText sRazonSocial
        ReporteA.Text7.SetText localConectado
        ReporteA.ReportTitle = sTitulo & sTexto
        frmEmite.CRViewer.ReportSource = ReporteA
    End If
End Sub


