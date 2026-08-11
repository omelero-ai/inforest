VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepRegistroVentaIntegrado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Ventas"
   ClientHeight    =   6045
   ClientLeft      =   2895
   ClientTop       =   2190
   ClientWidth     =   12405
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepRegistroVentaIntegrado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Exportar XML"
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
      Index           =   4
      Left            =   6240
      Picture         =   "frmRepRegistroVentaIntegrado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5400
      UseMaskColor    =   -1  'True
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
      Left            =   1860
      Picture         =   "frmRepRegistroVentaIntegrado.frx":10EC
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5400
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
      Left            =   3315
      Picture         =   "frmRepRegistroVentaIntegrado.frx":161E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton CmdActualiza 
      Caption         =   "Actualizar"
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
      Left            =   7650
      Picture         =   "frmRepRegistroVentaIntegrado.frx":1B50
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   1455
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
      Index           =   2
      Left            =   4770
      Picture         =   "frmRepRegistroVentaIntegrado.frx":1C18
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5400
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
      Height          =   5310
      Left            =   45
      TabIndex        =   21
      Top             =   45
      Width           =   12330
      Begin VB.CheckBox chkRegistroVenta 
         Caption         =   "Incluir comprobantes que no pertenescan al Registro de Ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   1920
         TabIndex        =   9
         Top             =   2640
         Width           =   5655
      End
      Begin VB.Frame Frame3 
         Caption         =   "Servidores En Linea"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   7680
         TabIndex        =   28
         Top             =   120
         Width           =   4575
         Begin VB.CheckBox chkTodo 
            Caption         =   "Seleccionar Todo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2880
            TabIndex        =   14
            Top             =   4560
            Width           =   1575
         End
         Begin MSComctlLib.ListView lstServidores 
            Height          =   4335
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   7646
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.ComboBox cboRedondeo 
         Height          =   360
         ItemData        =   "frmRepRegistroVentaIntegrado.frx":243A
         Left            =   1965
         List            =   "frmRepRegistroVentaIntegrado.frx":2444
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2175
         Width           =   1860
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
         Height          =   1575
         Left            =   1920
         TabIndex        =   10
         Top             =   3480
         Width           =   3375
         Begin VB.OptionButton optOpcion 
            Caption         =   "Agrupado Por Tipo Documento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   210
            TabIndex        =   29
            Top             =   1200
            Width           =   3075
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Por Local"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   210
            TabIndex        =   12
            Top             =   780
            Value           =   -1  'True
            Width           =   2835
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Por Sector"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   210
            TabIndex        =   11
            Top             =   390
            Width           =   3015
         End
      End
      Begin VB.CheckBox chkEstado 
         Caption         =   "Todos los Estados"
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
         Left            =   5400
         TabIndex        =   7
         Top             =   1725
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkTipoDocumento 
         Caption         =   "Todos los Documentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5400
         TabIndex        =   5
         Top             =   1350
         Value           =   1  'Checked
         Width           =   2310
      End
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   0
         Top             =   4440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "Xls"
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   330
         Left            =   1980
         TabIndex        =   2
         Top             =   870
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         _Version        =   393216
         Format          =   56623105
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   330
         Left            =   1980
         TabIndex        =   0
         Top             =   465
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         _Version        =   393216
         Format          =   56623105
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   330
         Left            =   4020
         TabIndex        =   1
         Top             =   465
         Width           =   1350
         _ExtentX        =   2381
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
         Format          =   56623107
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   330
         Left            =   4020
         TabIndex        =   3
         Top             =   870
         Width           =   1350
         _ExtentX        =   2381
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
         Format          =   56623107
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboTipoDocumento 
         Height          =   315
         Left            =   1980
         TabIndex        =   4
         Top             =   1305
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSDataListLib.DataCombo cboEstado 
         Height          =   315
         Left            =   1980
         TabIndex        =   6
         Top             =   1710
         Width           =   3375
         _ExtentX        =   5953
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
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Redondeo :"
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
         Left            =   855
         TabIndex        =   27
         Top             =   2265
         Width           =   1005
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Documento :"
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
         Left            =   105
         TabIndex        =   26
         Top             =   1770
         Width           =   1740
      End
      Begin VB.Label Label 
         Caption         =   "( Factura = Nota de Crédito )"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   5760
         TabIndex        =   25
         Top             =   1545
         Width           =   1815
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Documento :"
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
         Left            =   315
         TabIndex        =   24
         Top             =   1365
         Width           =   1530
      End
      Begin VB.Label Label 
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
         Index           =   1
         Left            =   720
         TabIndex        =   23
         Top             =   945
         Width           =   1125
      End
      Begin VB.Label Label 
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
         Index           =   0
         Left            =   615
         TabIndex        =   22
         Top             =   540
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
      Left            =   9090
      Picture         =   "frmRepRegistroVentaIntegrado.frx":2466
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5400
      Width           =   1455
   End
End
Attribute VB_Name = "frmRepRegistroVentaIntegrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCompania As Recordset
Dim RsCliente As Recordset
Dim RsEstado As Recordset
Dim RsTipoDocumento As Recordset
Dim ConsolidadoSector As New dsrRegistroVentaIntegradoS
Dim Consolidado As New dsrRegistroVentaIntegradoL
Dim Especial As New dsrRegistroVentaIntegradoE
Dim Identidad As String
Dim RsPrinter As Recordset
Dim sReporte As String
Dim sNombre As String
Dim sNombre1 As String
Dim sCriterio As String
Dim xCriterio As String
Dim sCliente As String
Dim fInicio As Date
Dim fFinal As Date

Sub LlenaCombos()
   With cboTipoDocumento
      'Compania
      Isql = "Select * from vTipoDocumento where Codigo<>'00' Order By Codigo"
      Set RsTipoDocumento = Lib.OpenRecordset(Isql, Cn)
      Set .RowSource = RsTipoDocumento
      .DataField = "Descripcion"
      .ListField = "Descripcion"
      .BoundColumn = "Codigo"
   End With
   
   With cboEstado
      'Estado
      Isql = "Select * from vEstadoDocumento where lActivo=1"
      Set RsEstado = Lib.OpenRecordset(Isql, Cn)
      Set .RowSource = RsEstado
      .DataField = "Descripcion"
      .ListField = "Descripcion"
      .BoundColumn = "Codigo"
   End With
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
  
    Cn.Execute "delete from " & sNombre
    Cn.Execute "delete from " & sNombre1
   If Index = 3 Then
      Unload Me
      Exit Sub
   End If
   Dim p, r As Integer
   sCriterio = ""
   xCriterio = ""
    r = 0
   p = 0
   For p = 1 To lstServidores.ListItems.Count
        If lstServidores.ListItems.Item(p).Checked = True Then
            r = r + 1
        End If
   Next p
   If r = 0 Then
      MsgBox "Debe seleccionar al menos un Item de la Lista de Servidores", vbCritical, sMensaje
   Exit Sub
   End If
   If dtpFecIni.value + dtpHorIni.value > dtpFecFin.value + dtpHorFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      Exit Sub
      dtpFecFin.SetFocus
   End If
      
   If chkTipoDocumento.value = 0 Then
      If cboTipoDocumento.Text = "" Then
         MsgBox "Debe escoger un Tipo de Documento", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " AND tTipoDocumento ='" & cboTipoDocumento.BoundText & "'"
      xCriterio = xCriterio & " AND Codigo ='" & cboTipoDocumento.BoundText & "'"
   End If
      
   If chkEstado.value = 0 Then
      If cboEstado.Text = "" Then
         MsgBox "Debe escoger un Estado del Documento", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " AND MDOCUMENTO.tEstadoDocumento ='" & cboEstado.BoundText & "'"
      xCriterio = xCriterio & " AND MNOTACREDITO.tEstadoDocumento ='" & cboEstado.BoundText & "'"
   End If
      
   cmdOpcion(Index).Enabled = False
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If RsPrinter.EOF = True Then
                  Screen.MousePointer = vbDefault
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  cmdOpcion(Index).Enabled = True
                  Exit Sub
               End If
               If optOpcion(0).value = True Then
                  ConsolidadoSector.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ConsolidadoSector.PaperOrientation = crLandscape
               ElseIf optOpcion(1).value = True Then
                  Consolidado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Consolidado.PaperOrientation = crLandscape
                ElseIf optOpcion(2).value = True Then
                  Especial.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Especial.PaperOrientation = crLandscape
               End If
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
          
          Case Is = 1 ' Imprimir
               Genera
               Screen.MousePointer = vbDefault
               If RsPrinter.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  cmdOpcion(Index).Enabled = True
                  Exit Sub
               End If
               If optOpcion(0).value = True Then
                  ConsolidadoSector.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ConsolidadoSector.PaperOrientation = crLandscape
                  ConsolidadoSector.PrintOut
               ElseIf optOpcion(1).value = True Then
                   Consolidado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                   Consolidado.PaperOrientation = crLandscape
                   Consolidado.PrintOut
                ElseIf optOpcion(2).value = True Then
                   Especial.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                   Especial.PaperOrientation = crLandscape
                   Especial.PrintOut
               End If
          
          Case Is = 2 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If RsPrinter.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  cmdOpcion(Index).Enabled = True
                  Exit Sub
               End If
               Screen.MousePointer = vbDefault
               'On Error GoTo Cancelar

               If optOpcion(0).value = True Then
                   ConsolidadoSector.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ConsolidadoSector.PaperOrientation = crLandscape
                  ConsolidadoSector.ExportOptions.FormatType = 21
                  ConsolidadoSector.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     cmdOpcion(Index).Enabled = True
                     Exit Sub
                  End If
                  ConsolidadoSector.ExportOptions.DiskFileName = cmdSave.FileName
                  ConsolidadoSector.Export False
               ElseIf optOpcion(1).value = True Then
                  Consolidado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Consolidado.PaperOrientation = crLandscape
                  Consolidado.ExportOptions.FormatType = 21
                  Consolidado.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     cmdOpcion(Index).Enabled = True
                     Exit Sub
                  End If
                  Consolidado.ExportOptions.DiskFileName = cmdSave.FileName
                  Consolidado.Export False
                ElseIf optOpcion(2).value = True Then
                  Especial.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  Especial.PaperOrientation = crLandscape
                  Especial.ExportOptions.FormatType = 21
                  Especial.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     cmdOpcion(Index).Enabled = True
                     Exit Sub
                  End If
                  Especial.ExportOptions.DiskFileName = cmdSave.FileName
                  Especial.Export False
               End If
          Case Is = 4 ' exportar xml
                 If MsgBox("¿Deseas generar archivo de exportacion de datos?", vbInformation + vbYesNo) = vbYes Then
                            Genera "XML"
                            If RsPrinter.EOF = True Then
                                       Screen.MousePointer = vbDefault
                                       MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                                       cmdOpcion(Index).Enabled = True
                                       Exit Sub
                             End If
                             cmdSave.Filter = "Archivos XML|*.xml"
                             cmdSave.ShowSave
                             If cmdSave.FileName = "" Then
                                cmdOpcion(Index).Enabled = True
                                Exit Sub
                             End If
                             Exportar RsPrinter, cmdSave.FileName
                             Screen.MousePointer = vbDefault
                Else
                     cmdOpcion(Index).Enabled = True
                    Exit Sub
                End If
          Case Is = 3 ' Salir
               Unload Me
   End Select
   cmdOpcion(Index).Enabled = True
Cancelar:
   Exit Sub
End Sub

'

Private Sub chkEstado_Click()
   If chkEstado.value = 1 Then
      cboEstado.Enabled = False
      cboEstado.Text = ""
   Else
      cboEstado.Enabled = True
   End If
End Sub

Private Sub chkTipoDocumento_Click()
   If chkTipoDocumento.value = 1 Then
      cboTipoDocumento.Enabled = False
      cboTipoDocumento.Text = ""
   Else
      cboTipoDocumento.Enabled = True
   End If
End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.value > dtpFecFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub Form_Load()
   Centrar Me
   LlenaCombos
   dtpFecIni.value = Date
   dtpFecFin.value = Date
  sNombre1 = dbTemporal(sCaja, 19, "codSectorlocal", "nvarchar(5)", "Sectorlocal", "nvarchar(250)", "tTipoDocumento", "nvarchar(5)", "tSerie", "nvarchar(5)", "tDocumento", "nvarchar(20)", "fregistro", "datetime", "tRuc", "nvarchar(15)", "tRazonSocial", "nvarchar(100)", "tDireccion", "nvarchar(200)", "nneto", "Float", "nimpuesto1", "Float", "nimpuesto2", "Float", "nimpuesto3", "Float", "nVenta", "Float", "nRecargo", "float", "ndescuento", "Float", "tEstado", "nvarchar(50)", "tCodigoEstado", "nvarchar(5)", "tCodigoCliente", "nvarchar(10)")
   sNombre = dbTemporal(sCaja, 12, "sector", "nvarchar(50)", "codlocal", "nvarchar(2)", "local", "nvarchar(50)", "sunat", "nvarchar(3)", "minimo", "nvarchar(15)", "maximo", "nvarchar(15)", "neto", "Float", "impuesto1", "Float", "impuesto2", "Float", "impuesto3", "Float", "Venta", "Float", "descuento", "Float")
   cboTipoDocumento.Enabled = False
   cboEstado.Enabled = False
   'configuraListaServidores
   obtieneListaServidores
End Sub
Public Sub Genera(Optional xml As String)
   Screen.MousePointer = vbHourglass
   On Error GoTo sigue
   Dim conectado As Boolean
   Dim oComando As clsComando
   Dim sOrden As String
   Dim RsDatos As New ADODB.Recordset
   Dim w As Integer
   Dim z As Long
   Dim cs As String
   Dim cl As String
   Dim sector As String
   Dim llocal As String
   Set oComando = New clsComando
   
   

   
   fInicio = Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "HH:mm")
   fFinal = Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(dtpHorFin.value, "HH:mm")
   
    For w = 1 To lstServidores.ListItems.Count
        If lstServidores.ListItems.Item(w).Checked = True Then
                conectado = validaConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))
                If conectado = False Then
                    MsgBox "Error Fatal: No se pudo conectar con  " & lstServidores.ListItems.Item(w).SubItems(2), vbCritical, sMensaje
                    lstServidores.ListItems.Item(w).Checked = False
                    GoTo sigue
                End If
                If Not oComando.CreateCmdSp("spRep_RegVentaIntegrado", devuelveConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))) Then
                    Set oComando = Nothing
                Exit Sub
                End If
                      oComando.CreateParameter "@flagTipo", adBoolean, adParamInput, 1, Me.optOpcion(2).value
                      oComando.CreateParameter "@flagRedondeo", adBoolean, adParamInput, 1, IIf(cboRedondeo.Text = "Dos Decimales", True, False)
                      oComando.CreateParameter "@flagRegventa", adBoolean, adParamInput, 1, IIf(Me.chkRegistroVenta.value = 1, False, True)
                      oComando.CreateParameter "@tTipoDoc", adVarChar, adParamInput, 20, cboTipoDocumento.BoundText
                      oComando.CreateParameter "@tEstadoDoc", adVarChar, adParamInput, 20, cboEstado.BoundText
                      oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
                      oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
                      
                      If Not oComando.GetParamOK Then
                            Set oComando = Nothing
                            Exit Sub
                      End If
                 
                    cl = lstServidores.ListItems.Item(w)
                    sector = lstServidores.ListItems.Item(w).SubItems(1)
                    llocal = lstServidores.ListItems.Item(w).SubItems(2)
                  Set RsDatos = oComando.GetSP()
                     If Not (RsDatos.EOF Or RsDatos.BOF) Then
                                    While Not RsDatos.EOF
                                            If optOpcion(2).value = True Then
                                                                                              'codSectorlocal,Sectorlocal,           tTipoDocumento,                 tSerie,             tDocumento,              fregistro,                 tRuc,                      tRazonSocial,                tDireccion,                 nneto,                  nimpuesto1,               nimpuesto2,                 nimpuesto3,                 nVenta,                 nRecargo,                 ndescuento,                 tEstado,                                     tCodigoEstado, tCodigoCliente
                                                Isql = "insert into " & sNombre1 & " values ('" & cl & "','" & llocal & "','" & RsDatos.Fields(0) & "','" & RsDatos.Fields(1) & "','" & RsDatos.Fields(2) & "','" & Format(RsDatos.Fields(3), "mm/dd/yyyy") & "','" & RsDatos.Fields(4) & "','" & RsDatos.Fields(5) & "','" & RsDatos.Fields(6) & "'," & RsDatos.Fields(7) & "," & RsDatos.Fields(8) & "," & RsDatos.Fields(9) & "," & RsDatos.Fields(10) & " ," & RsDatos.Fields(11) & "," & RsDatos.Fields(12) & "," & RsDatos.Fields(13) & ",'" & RsDatos.Fields(14) & "','" & RsDatos.Fields(15) & "','" & RsDatos.Fields(12) & "'  )"
                                            Else
                                                Isql = "insert into " & sNombre & " values ('" & sector & "','" & cl & "','" & llocal & "','" & RsDatos.Fields(0) & "','" & RsDatos.Fields(1) & "','" & RsDatos.Fields(2) & "'," & RsDatos.Fields(3) & "," & RsDatos.Fields(4) & "," & RsDatos.Fields(5) & "," & RsDatos.Fields(6) & "," & RsDatos.Fields(7) & "," & RsDatos.Fields(8) & " )"
                                                '  Debug.Print Isql
                                            End If
                                            Cn.Execute Isql
                                            RsDatos.MoveNext
                                     Wend
                                
                    End If
        End If
sigue:
    Next w
      
                     If optOpcion(0).value = True Then ' por sector
                            Isql = "select sector,sunat,min(minimo),max(maximo),sum(neto),sum(impuesto1),sum(impuesto2),sum(impuesto3),sum(venta),sum(descuento) from " & sNombre & " group by  sector,sunat "
                            
                                  Set RsPrinter = Lib.OpenRecordset(Isql, Cn)
                                  If UCase(xml) = UCase("xml") Then: Exit Sub
                                  ConsolidadoSector.DiscardSavedData
                                  ConsolidadoSector.Database.SetDataSource RsPrinter
                                  ConsolidadoSector.ReportTitle = "Agrupado por Sectores Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy")
                                  ConsolidadoSector.Text10.SetText IIf(sImpuesto1 = "", "No definido", sImpuesto1)
                                  ConsolidadoSector.Text14.SetText IIf(sImpuesto2 = "", "No definido", sImpuesto2)
                                  ConsolidadoSector.Text13.SetText IIf(sImpuesto3 = "", "No definido", sImpuesto3)
                                  ConsolidadoSector.Text18.SetText "Tipo Documento: " + IIf(Me.chkTipoDocumento.value = 0, Me.cboTipoDocumento.Text, "TODOS") + " / Estado: " + IIf(Me.chkEstado.value = 0, Me.cboEstado.Text, "TODOS") + " (Monto Expresado en Soles)"
                                  ConsolidadoSector.Text8.SetText sRazonSocial
                                  frmEmite.CRViewer.DisplayGroupTree = False
                                  frmEmite.CRViewer.ReportSource = ConsolidadoSector
                     ElseIf optOpcion(1).value = True Then
                            Isql = "select Sector + ' :  ' + Local,sector,codlocal,local,sunat,minimo,maximo, neto,impuesto1,impuesto2,impuesto3,venta,descuento  from " & sNombre & " order by 3,4,5 "
                                  Set RsPrinter = Lib.OpenRecordset(Isql, Cn)
                                  If UCase(xml) = UCase("xml") Then: Exit Sub
                                  Consolidado.DiscardSavedData
                                  Consolidado.Database.SetDataSource RsPrinter
                                  Consolidado.ReportTitle = "Agrupado por Locales Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy")
                                  Consolidado.Text10.SetText IIf(sImpuesto1 = "", "No definido", sImpuesto1)
                                  Consolidado.Text14.SetText IIf(sImpuesto2 = "", "No definido", sImpuesto2)
                                  Consolidado.Text13.SetText IIf(sImpuesto3 = "", "No definido", sImpuesto3)
                                  Consolidado.Text18.SetText "Tipo Documento: " + IIf(Me.chkTipoDocumento.value = 0, Me.cboTipoDocumento.Text, "TODOS") + " / Estado: " + IIf(Me.chkEstado.value = 0, Me.cboEstado.Text, "TODOS") + " (Monto Expresado en Soles)"
                                  Consolidado.Text8.SetText sRazonSocial
                                  frmEmite.CRViewer.DisplayGroupTree = False
                                  frmEmite.CRViewer.ReportSource = Consolidado
                    ElseIf optOpcion(2).value = True Then
                                  'Isql = "select Sector + ' :  ' + Local,sector,codlocal,local,sunat,minimo,maximo, neto,impuesto1,impuesto2,impuesto3,venta,descuento  from " & sNombre & " order by 3,4,5 "
                                  Isql = " select codSectorlocal,Sectorlocal,tTipoDocumento,tSerie,tDocumento, fregistro ,tRuc,tRazonSocial,tDireccion,nneto,nimpuesto1,nimpuesto2,nimpuesto3,                 nVenta,                 nRecargo,                 ndescuento,                 tEstado,                                     tCodigoEstado, tCodigoCliente from " & sNombre1 & " order by codSectorLocal,ttipoDocumento,tSerie "
                                  Set RsPrinter = Lib.OpenRecordset(Isql, Cn)
                                  If UCase(xml) = UCase("xml") Then: Exit Sub
                                  Especial.DiscardSavedData
                                  Especial.Database.SetDataSource RsPrinter
                                  Especial.ReportTitle = "Agrupado por Locales Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy")
                                  Especial.Text16.SetText IIf(sImpuesto1 = "", "No definido", sImpuesto1)
                                  Especial.Text17.SetText IIf(sImpuesto2 = "", "No definido", sImpuesto2)
                                  Especial.Text18.SetText IIf(sImpuesto3 = "", "No definido", sImpuesto3)
                                  Especial.Text20.SetText "Tipo Documento: " + IIf(Me.chkTipoDocumento.value = 0, Me.cboTipoDocumento.Text, "TODOS") + " / Estado: " + IIf(Me.chkEstado.value = 0, Me.cboEstado.Text, "TODOS") + " (Monto Expresado en Soles)"
                                  Especial.Text8.SetText sRazonSocial
                                  frmEmite.CRViewer.DisplayGroupTree = False
                                  frmEmite.CRViewer.ReportSource = Especial
                    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsPrinter = Nothing
   Set frmRepRegistroVentaIntegrado = Nothing
    Set RsCliente = Nothing
   Set RsEstado = Nothing
   Set RsTipoDocumento = Nothing
   If sNombre <> "" Then
      Cn.Execute "Drop Table " & sNombre
   End If
   If sNombre1 <> "" Then
      Cn.Execute "Drop Table " & sNombre1
   End If
End Sub

Private Sub obtieneListaServidores()
 Dim rsS As New ADODB.Recordset
 Set rsS = rsListaServidores
 Me.lstServidores.Checkboxes = True
 lstServidores.ColumnHeaders.Clear
 Me.lstServidores.ListItems.Clear
    Dim Item As ListItem
    With lstServidores
        .ColumnHeaders.Add , , " Código ", 800
        .ColumnHeaders.Add , , " Sector ", 1100
         .ColumnHeaders.Add , , " Local ", 1100
         .ColumnHeaders.Add , , " Estado ", 1100
         .ColumnHeaders.Add , , "Ip", 0
         .ColumnHeaders.Add , , "Bd", 0
        .View = lvwReport
        .GridLines = True
        
     Dim X As Integer
        rsS.MoveFirst
        For X = 0 To rsS.RecordCount - 1
            Set Item = .ListItems.Add(, , rsS.Fields(0))
            Item.SubItems(1) = rsS.Fields(1)
            Item.SubItems(2) = rsS.Fields(2)
            Item.SubItems(3) = rsS.Fields(5)
            Item.SubItems(4) = rsS.Fields(3)
            Item.SubItems(5) = rsS.Fields(4)
        rsS.MoveNext
        Next X
            End With

End Sub
Private Sub configuraListaServidores()
 Dim rsS As New ADODB.Recordset
 Set rsS = devuelveServidoresConectados
 Me.lstServidores.Checkboxes = True
 lstServidores.ColumnHeaders.Clear
 Me.lstServidores.ListItems.Clear
    Dim Item As ListItem
    With lstServidores
        .ColumnHeaders.Add , , " Código ", 800
        .ColumnHeaders.Add , , " Sector ", 1100
         .ColumnHeaders.Add , , " Local ", 1100
         .ColumnHeaders.Add , , " Estado ", 1100
         .ColumnHeaders.Add , , "Ip", 0
         .ColumnHeaders.Add , , "Bd", 0
        .View = lvwReport
        .GridLines = True
        
     Dim X As Integer
        rsS.MoveFirst
        For X = 0 To rsS.RecordCount - 1
            Set Item = .ListItems.Add(, , rsS.Fields(0))
            Item.SubItems(1) = rsS.Fields(1)
            Item.SubItems(2) = rsS.Fields(2)
            Item.SubItems(3) = rsS.Fields(5)
            Item.SubItems(4) = rsS.Fields(3)
            Item.SubItems(5) = rsS.Fields(4)
        rsS.MoveNext
        Next X
            End With
End Sub

Private Sub lstServidores_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.SubItems(3) = "Desconectado" Then
        MsgBox "El enlace con el Servidor Seleccionado no es Posible", vbCritical, sMensaje
        Item.Checked = False
    End If
End Sub
Private Sub CmdActualiza_Click()
If MsgBox("Esta operación puede tardar varios minutos" & vbCrLf & "¿Desea continuar?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
    Screen.MousePointer = vbHourglass
    configuraListaServidores
    Me.chkTodo.value = 0
    MsgBox "Lista De Servidores Actualizada", vbInformation, sMensaje
    Screen.MousePointer = vbDefault
End If
End Sub

Public Sub activaCheck(Tipo As Boolean)
    Dim o As Integer
    For o = 1 To lstServidores.ListItems.Count
        If Tipo = True Then
            If lstServidores.ListItems.Item(o).SubItems(3) <> "Desconectado" Then

                lstServidores.ListItems.Item(o).Checked = Tipo
            End If
        Else
            lstServidores.ListItems.Item(o).Checked = Tipo
        End If
    Next o
End Sub

Private Sub chkTodo_Click()
        activaCheck IIf(chkTodo.value = 1, True, False)
End Sub


'XML

Sub Exportar(rsReporte As ADODB.Recordset, ruta As String)
    On Error GoTo err:
        Dim clasexml As clsxml
        Set clasexml = New clsxml
        If Not clasexml.CrearXML(rsReporte, ruta) Then
           Set clasexml = Nothing
        End If
        Exit Sub
err:
    MsgBox "Error al intentar generar el archivo XML", vbCritical, sMensaje
        
End Sub

