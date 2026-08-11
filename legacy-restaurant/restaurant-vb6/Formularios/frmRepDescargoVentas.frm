VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepDescargoVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descargo de Ventas"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
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
      Left            =   1080
      Picture         =   "frmRepDescargoVentas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3960
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
      Left            =   2550
      Picture         =   "frmRepDescargoVentas.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3960
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
      Left            =   5400
      Picture         =   "frmRepDescargoVentas.frx":0A64
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3960
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
      Height          =   3885
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7395
      Begin VB.Frame Frame3 
         Caption         =   "Tipo Reporte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1680
         TabIndex        =   28
         Top             =   2760
         Width           =   3255
         Begin VB.OptionButton OptResumido 
            Caption         =   "Resumido(Pedido)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   600
            Width           =   2535
         End
         Begin VB.OptionButton optDetallado 
            Caption         =   "Detallado (Pedido - Item)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   300
            Value           =   -1  'True
            Width           =   2655
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   5040
         TabIndex        =   24
         Top             =   1800
         Width           =   1815
         Begin VB.OptionButton optNoDes 
            Caption         =   "No Descargados"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton optDes 
            Caption         =   "Descargados"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton optTodos 
            Caption         =   "Todos"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.CheckBox chkProducto 
         Caption         =   "Todos los Insumos"
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
         Left            =   5040
         TabIndex        =   7
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2115
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
         Index           =   0
         Left            =   4140
         Picture         =   "frmRepDescargoVentas.frx":0B56
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1035
         Width           =   765
      End
      Begin VB.TextBox txtInsumo 
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   2370
      End
      Begin VB.CheckBox chkSubGrupo 
         Caption         =   "Todos los Sub Grupos"
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
         Left            =   5040
         TabIndex        =   4
         Top             =   675
         Value           =   1  'Checked
         Width           =   2340
      End
      Begin VB.CheckBox chkGrupo 
         Caption         =   "Todas los Grupo"
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
         Left            =   5040
         TabIndex        =   3
         Top             =   270
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin VB.CheckBox chkArea 
         Caption         =   "Todas las Areas"
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
         Left            =   5040
         TabIndex        =   2
         Top             =   1485
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1680
         TabIndex        =   8
         Top             =   2295
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
         Format          =   57081857
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   1680
         TabIndex        =   9
         Top             =   1890
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
         Format          =   57081857
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   3675
         TabIndex        =   10
         Top             =   1890
         Width           =   1230
         _ExtentX        =   2170
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
         Format          =   57081859
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   3675
         TabIndex        =   11
         Top             =   2295
         Width           =   1230
         _ExtentX        =   2170
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
         Format          =   57081859
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboGrupo 
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Top             =   270
         Width           =   3225
         _ExtentX        =   5689
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
      Begin MSDataListLib.DataCombo cboSubGrupo 
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   675
         Width           =   3225
         _ExtentX        =   5689
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
      Begin MSDataListLib.DataCombo cboArea 
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Top             =   1485
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   20
         Top             =   1890
         Width           =   1545
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   19
         Top             =   2295
         Width           =   1545
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insumos :"
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
         Index           =   7
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Grupo:"
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
         Index           =   6
         Left            =   690
         TabIndex        =   17
         Top             =   675
         Width           =   975
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo :"
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
         Index           =   4
         Left            =   1020
         TabIndex        =   16
         Top             =   270
         Width           =   645
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Area Producción :"
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
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   1485
         Width           =   1545
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
      Left            =   3960
      Picture         =   "frmRepDescargoVentas.frx":0C58
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   240
      Top             =   2895
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepDescargoVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNombre     As String
Dim rsReporte   As Recordset
Dim RsArea      As Recordset
Dim RsProducto  As Recordset

Dim dsrReporte   As New dsrRepDescargo
Dim dsrReporteResumido   As New dsrRepDescargoResumido
'Dim Resumido    As New dsrInsumosR
'Dim ControlDiario    As New dsrInsumosCD

Dim sCriterio   As String
Dim sInsumo     As String
Dim sFiltro     As String
Dim sTitulo     As String
Dim sPrecio     As String
Dim sTexto      As String

Dim fInicio As Date
Dim fFinal As Date
Dim familia As String
Dim subFamilia As String
Dim Area As String

Dim RsGrupo As Recordset
Dim RsSubgrupo As Recordset

Sub LlenaCombos()
    With cboGrupo
         Isql = "Select * from vGrupo order by Descripcion"
         Set RsGrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsGrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
           
    With cboSubGrupo
         Isql = "Select * from vSubgrupo order by Descripcion"
         Set RsSubgrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubgrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboArea
         Isql = "Select * from vArea"
         Set RsArea = Lib.OpenRecordset(Isql, CnAlmacen)
         Set .RowSource = RsArea
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
End Sub

Private Sub cboGrupo_Click(Area As Integer)
    cboSubGrupo.BoundText = ""
    With cboSubGrupo
         Isql = "Select * from vsubgrupo " & IIf(cboGrupo.BoundText = "", "", "where substring(Codigo,1,2) = '" & cboGrupo.BoundText & "'") & " order by Descripcion "
         Set RsSubgrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubgrupo
    End With
End Sub

Private Sub chkArea_Click()
    If chkArea.Value = 1 Then
      cboArea.Enabled = False
      cboArea.Text = ""
   Else
      cboArea.Enabled = True
   End If
End Sub

Private Sub chkGrupo_Click()
   If chkGrupo.Value = 1 Then
      cboGrupo.Enabled = False
      cboGrupo.Text = ""
   Else
      cboGrupo.Enabled = True
   End If
End Sub

Private Sub chkSubGrupo_Click()
   If chkSubGrupo.Value = 1 Then
      cboSubGrupo.Enabled = False
      cboSubGrupo.Text = ""
   Else
      cboSubGrupo.Enabled = True
   End If
End Sub
Private Sub chkproducto_Click()
   If chkProducto.Value = 1 Then
      sInsumo = ""
      txtInsumo.Text = ""
      cmdBusca(0).Enabled = False
   Else
      cmdBusca(0).Enabled = True
   End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
On Error GoTo fin
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   Set dsrReporte = New dsrRepDescargo
   
   familia = ""
   subFamilia = ""

   Area = ""
   sCriterio = ""
   sTitulo = ""
   sFiltro = ""
   
   If chkGrupo.Value = 0 Then
      If cboGrupo.Text = "" Then
         MsgBox "Debe escoger la Familia", vbCritical, sMensaje
         Exit Sub
      End If
        familia = cboGrupo.Text
   End If
    
   If chkSubGrupo.Value = 0 Then
      If cboSubGrupo.Text = "" Then
         MsgBox "Debe escoger la sub familia", vbCritical, sMensaje
         Exit Sub
      End If
        subFamilia = cboSubGrupo.Text
   End If
   If chkProducto.Value = 0 Then
      If sInsumo = "" Then
         MsgBox "Debe escoger el producto", vbCritical, sMensaje
         Exit Sub
      End If
   End If
   
   If chkArea.Value = 0 Then
    If Me.cboArea.Text = "" Then
        MsgBox "Debe escoger el area", vbCritical, sMensaje
        Exit Sub
    End If
    'CMiranda------------------------------------------------------------------------------
    Area = Me.cboArea.BoundText
    'Fin CMiranda--------------------------------------------------------------------------
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
            If optDetallado Then
                dsrReporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                dsrReporte.PaperOrientation = crLandscape
            End If
            If OptResumido Then
                dsrReporteResumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                dsrReporteResumido.PaperOrientation = crLandscape
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
            If optDetallado Then
                dsrReporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                dsrReporte.PaperOrientation = crLandscape
                dsrReporte.PrintOut
            End If
            If OptResumido Then
                dsrReporteResumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                dsrReporteResumido.PaperOrientation = crLandscape
                dsrReporteResumido.PrintOut
            End If
            
                
          Case Is = 2 ' Salir
               Unload Me
          
          Case Is = 3 ' Exportar
            Genera
            If rsReporte.EOF = True Then
                MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                Exit Sub
            End If
            ExportaExcel rsReporte

   End Select
   Exit Sub
fin:
          MsgBox "Mensaje: " & error, vbInformation
End Sub

Private Sub cmdBusca_Click(Index As Integer)
   Dim xCriterio As String
   Select Case Index
          Case Is = 0
               xCriterio = "lActivo = 1 "
               Isql = "select tCodigoProducto as Codigo, tDetallado as Descripcion from vProducto where " & xCriterio & " order by Descripcion"
               frmBuscaAlmacen.nPredeterm = 1
               Call ConfGrilla(2, frmBuscaAlmacen.grdGrilla, "Codigo", 2, "Codigo", 2300, 2, 0, "", _
                                                      "Insumo", 2, "Descripcion", 5000, 0, 0, "")
               frmBuscaAlmacen.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
               sInsumo = sCodigo
               txtInsumo.Text = sDescrip
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
    cmdBusca(0).Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
End Sub

Public Sub Genera()
On Error GoTo fin

    Dim oComando As clsComando
    Dim fInicio As String
    Dim fFinal As String
    
    fInicio = Format(dtpFecIni.Value, "yyyyMMdd") & " " & Format(Me.dtpHorIni.Value, "HH:mm")
    fFinal = Format(dtpFecFin.Value, "yyyyMMdd") & " " & Format(Me.dtpHorFin.Value, "HH:mm")
    
    Screen.MousePointer = vbHourglass
    Set oComando = New clsComando
  
    If Not oComando.CreateCmdSp("usp_RepInforest_DescargoVenta", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    
    If BDLink = "1" Then
        oComando.CreateParameter "@Almacen", adVarChar, adParamInput, 50, "[" & sAlmacenRuta & "].[" & sAlmacenMDB & "]" 'CASO CORP.FERNANDEZ
    Else
        oComando.CreateParameter "@Almacen", adVarChar, adParamInput, 50, sAlmacenMDB   'CASO CORP.FERNANDEZ
    End If
    oComando.CreateParameter "@FechaIni", adVarChar, adParamInput, 50, fInicio
    oComando.CreateParameter "@FechaFin", adVarChar, adParamInput, 50, fFinal
    oComando.CreateParameter "@sTemporal", adVarChar, adParamInput, 50, ""
    oComando.CreateParameter "@Local", adVarChar, adParamInput, 50, sLocal
    oComando.CreateParameter "@Grupo", adVarChar, adParamInput, 50, Me.cboGrupo.BoundText
    oComando.CreateParameter "@SubGrupo", adVarChar, adParamInput, 50, Me.cboSubGrupo.BoundText
    oComando.CreateParameter "@Insumo", adVarChar, adParamInput, 50, sInsumo
    oComando.CreateParameter "@Area", adVarChar, adParamInput, 50, Me.cboArea.BoundText
    oComando.CreateParameter "@Descargo", adVarChar, adParamInput, 50, IIf(Me.optTodos.Value, "T", IIf(Me.optDes.Value, "D", "ND"))
    oComando.CreateParameter "@tipooper", adInteger, adParamInput, 10, IIf(optDetallado.Value, 1, 2)
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set rsReporte = oComando.GetSP()
    If Me.optDetallado Then
        dsrReporte.DiscardSavedData
        dsrReporte.Database.SetDataSource rsReporte
        If Me.optTodos Then
            dsrReporte.Text27.SetText "TODOS"
        ElseIf Me.optDes Then
            dsrReporte.Text27.SetText "DESCARGADOS"
        Else
            dsrReporte.Text27.SetText "NO DESCARGADOS"
        End If

        dsrReporte.Text8.SetText sRazonSocial
        frmEmite.CRViewer.ReportSource = dsrReporte
    End If
    If Me.OptResumido Then
        dsrReporteResumido.DiscardSavedData
        dsrReporteResumido.Database.SetDataSource rsReporte
        If Me.optTodos Then
            dsrReporteResumido.Text5.SetText "TODOS"
        ElseIf Me.optDes Then
            dsrReporteResumido.Text5.SetText "DESCARGADOS"
        Else
            dsrReporteResumido.Text5.SetText "NO DESCARGADOS"
        End If
        
        dsrReporteResumido.Text8.SetText sRazonSocial
        frmEmite.CRViewer.ReportSource = dsrReporteResumido
    End If

    Screen.MousePointer = vbDefault
Exit Sub
fin:
    MsgBox "Mensaje: " & error, vbInformation
End Sub





