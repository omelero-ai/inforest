VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmGenerarSunat 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar TXT SUNAT"
   ClientHeight    =   4905
   ClientLeft      =   2895
   ClientTop       =   2190
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGenerarsunat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
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
      Height          =   4065
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9090
      Begin VB.ComboBox cboOrden 
         Height          =   360
         ItemData        =   "frmGenerarsunat.frx":000C
         Left            =   2520
         List            =   "frmGenerarsunat.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2340
         Width           =   1860
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
         Left            =   6465
         TabIndex        =   9
         Top             =   630
         Value           =   1  'Checked
         Width           =   2325
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
         Left            =   6465
         TabIndex        =   8
         Top             =   1005
         Value           =   1  'Checked
         Width           =   2325
      End
      Begin VB.CheckBox chkCliente 
         Caption         =   "Todos los Clientes"
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
         Left            =   6465
         TabIndex        =   7
         Top             =   210
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
         Left            =   5460
         Picture         =   "frmGenerarsunat.frx":003A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   165
         Width           =   765
      End
      Begin VB.TextBox txtCliente 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   195
         Width           =   2775
      End
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   7170
         Top             =   -510
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "Xls"
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   330
         Left            =   2520
         TabIndex        =   11
         Top             =   3480
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         _Version        =   393216
         Format          =   63963137
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   330
         Left            =   2520
         TabIndex        =   12
         Top             =   2925
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         _Version        =   393216
         Format          =   63963137
         CurrentDate     =   37539
      End
      Begin MSDataListLib.DataCombo cboTipoDocumento 
         Height          =   315
         Left            =   2520
         TabIndex        =   13
         Top             =   585
         Width           =   3615
         _ExtentX        =   6376
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
         Left            =   2520
         TabIndex        =   14
         Top             =   990
         Width           =   3615
         _ExtentX        =   6376
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
      Begin MSDataListLib.DataCombo cboCaja 
         Height          =   315
         Left            =   2520
         TabIndex        =   15
         Top             =   1455
         Width           =   3615
         _ExtentX        =   6376
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
      Begin MSDataListLib.DataCombo cboMotivoSunat 
         Height          =   315
         Left            =   2520
         TabIndex        =   24
         Top             =   1920
         Width           =   3615
         _ExtentX        =   6376
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
         Caption         =   "Motivo :"
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
         Left            =   1680
         TabIndex        =   25
         Top             =   1965
         Width           =   705
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   1155
         TabIndex        =   23
         Top             =   2415
         Width           =   1290
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
         Left            =   1095
         TabIndex        =   22
         Top             =   2985
         Width           =   1230
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
         Left            =   1200
         TabIndex        =   21
         Top             =   3525
         Width           =   1125
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
         Left            =   915
         TabIndex        =   20
         Top             =   645
         Width           =   1530
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
         Left            =   6855
         TabIndex        =   19
         Top             =   825
         Width           =   1815
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
         Left            =   705
         TabIndex        =   18
         Top             =   1050
         Width           =   1740
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente :"
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
         Index           =   10
         Left            =   1725
         TabIndex        =   17
         Top             =   255
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caja :"
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
         Index           =   11
         Left            =   1920
         TabIndex        =   16
         Top             =   1500
         Width           =   510
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Exportar TXT"
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
      Left            =   6840
      Picture         =   "frmGenerarsunat.frx":013C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
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
      Left            =   4440
      Picture         =   "frmGenerarsunat.frx":095E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
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
      Left            =   2895
      Picture         =   "frmGenerarsunat.frx":0A50
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4215
      Visible         =   0   'False
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
      Left            =   1440
      Picture         =   "frmGenerarsunat.frx":0F82
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4215
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmGenerarSunat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsCliente As Recordset
Dim RsEstado As Recordset
Dim RsTipoDocumento As Recordset
Dim RsTipoPago As Recordset
Dim RsMotivoSunat As Recordset



Dim Identidad As String
Dim RsPrinter As Recordset
Dim RsCaja    As Recordset
Dim sReporte As String
Dim sCriterio As String
Dim xCriterio As String
Dim vCriterio As String
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
   
    With cboCaja
         Isql = "Select tCaja as Codigo, tDescripcion as Descripcion from tCaja order by 1"
         Set RsCaja = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsCaja
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboMotivoSunat
         Isql = "Select TCODIGO as Codigo, TCODIGO+'-'+tDetallado as Descripcion from ttabla where ttabla='MOTIVOSUNAT' order by 1"
         Set RsMotivoSunat = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMotivoSunat
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    

   
End Sub


'Private Sub ChkCaja_Click()
'   If ChkCaja.value = 1 Then
'      cboCaja.Enabled = False
'      cboCaja.Text = ""
'   Else
'      cboCaja.Enabled = True
'   End If
'End Sub



Private Sub cmdBusca_Click()
   Dim xCriterio As String
   Isql = "Select tCodigoCliente as Codigo, tIdentidad as Identidad, tEmpresa as Descripcion from TCLIENTE order by tEmpresa"
   frmBusca.cboCriterio.Enabled = True
   frmBusca.nPredeterm = 2
   Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                          "Identificador", 2, "Identidad", 1500, 0, 0, "", _
                                          "Razón Comercial", 2, "Descripcion", 4500, 0, 0, "")
   frmBusca.Show vbModal
   If Not wEnter Then
      Exit Sub
   End If
   'sCliente = Calcular("Select tCodigoCliente As Codigo From tCliente Where tIdentidad ='" & sCodigo & "'", Cn)
   sCliente = sCodigo
   txtCliente.Text = sDescrip
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   
   If Index = 3 Then
      Unload Me
      Exit Sub
   End If
   
   sCriterio = ""
   xCriterio = ""
   vCriterio = ""

'DiaContable

   If dtpFecIni.Value > dtpFecFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      Exit Sub
      dtpFecFin.SetFocus
   End If
      
   If chkCliente.Value = 0 Then
      If sCliente = "" Then
         MsgBox "Debe escoger el Cliente", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " AND tCodigoCliente ='" & sCliente & "'"
      xCriterio = xCriterio & " AND tCodigoCliente ='" & sCliente & "'"
      vCriterio = vCriterio & " AND CodigoCliente ='" & sCliente & "'"                          'jesus 160512
   End If
   
   If chkTipoDocumento.Value = 0 Then
      If cboTipoDocumento.Text = "" Then
         MsgBox "Debe escoger un Tipo de Documento", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " AND tTipoDocumento ='" & cboTipoDocumento.BoundText & "'"
      xCriterio = xCriterio & " AND Codigo ='" & cboTipoDocumento.BoundText & "'"
      vCriterio = vCriterio & " AND CodigoDocumento ='" & cboTipoDocumento.BoundText & "'"      'jesus 160512
   End If
      
   If chkEstado.Value = 0 Then
      If cboEstado.Text = "" Then
         MsgBox "Debe escoger un Estado del Documento", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " AND MDOCUMENTO.tEstadoDocumento ='" & cboEstado.BoundText & "'"
      xCriterio = xCriterio & " AND MNOTACREDITO.tEstadoDocumento ='" & cboEstado.BoundText & "'"
      vCriterio = vCriterio & " AND EstadoDocumento ='" & cboEstado.BoundText & "'"             'jesus 160512
   End If
      
   
      If cboCaja.Text = "" Then
         MsgBox "Debe escoger una Caja", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & " AND MDOCUMENTO.tcaja ='" & cboCaja.BoundText & "'"
      xCriterio = xCriterio & " AND MNOTACREDITO.tcaja ='" & cboCaja.BoundText & "'"
      vCriterio = vCriterio & " AND Caja ='" & cboCaja.BoundText & "'"                          'jesus 160512
   
   If cboMotivoSunat.Text = "" Then
         MsgBox "Debe escoger una Motivo", vbCritical, sMensaje
         Exit Sub
    End If
      sCriterio = sCriterio & " AND MDOCUMENTO.tcaja ='" & cboCaja.BoundText & "'"
      xCriterio = xCriterio & " AND MNOTACREDITO.tcaja ='" & cboCaja.BoundText & "'"
      vCriterio = vCriterio & " AND Caja ='" & cboCaja.BoundText & "'"                          'jesus 160512
   
   cmdOpcion(Index).Enabled = False
   Select Case Index
   
         ' Case Is = 0 ' Preview
                
          '                  Genera1
                
           '    If RsPrinter.EOF = True Then
             '     Screen.MousePointer = vbDefault
            '      MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
             '     cmdOpcion(Index).Enabled = True
             '     Exit Sub
             '  End If
               
               
             '     RsSunat.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
             '     RsSunat.PaperOrientation = crLandscape
               
             '  frmEmite.CRViewer.ViewReport
             '  frmEmite.Show vbModal
          
          'Case Is = 1 ' Imprimir
          
                 
         '                    Genera1
                 
          '     Screen.MousePointer = vbDefault
            '   If RsPrinter.EOF = True Then
           '       MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
             '     cmdOpcion(Index).Enabled = True
              '    Exit Sub
              ' End If
                  
               '   RsSunat.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
              '    RsSunat.PaperOrientation = crLandscape
               '   RsSunat.PrintOut
               
               
               
          
          Case Is = 2 ' Exportar
          
                             Genera1
               
                
               Screen.MousePointer = vbDefault
               If RsPrinter.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  cmdOpcion(Index).Enabled = True
                  Exit Sub
               End If
               Screen.MousePointer = vbDefault
               'On Error GoTo Cancelar

                  'RsSunat.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  'RsSunat.PaperOrientation = crLandscape
                  
                  'EXPORTAR TXT VER CODIGO
                  'RsSunat.ExportOptions.FormatType = exportformattype.
                  'RsSunat.ExportOptions.DestinationType = 1
                  
                  cmdSave.Filter = "Archivo de texto|*.txt"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     cmdOpcion(Index).Enabled = True
                     Exit Sub
                  End If
                   Dim intFile As Integer
                    Dim strFile As String
                  
                  RsPrinter.MoveFirst
                  Open cmdSave.FileName For Output As #1
                  
                    'strFile = "c:\temp\file.txt" 'the file you want to save to
                    'intFile = FreeFile
                    
                  Do While Not RsPrinter.EOF
                  
                    Print #1, Trim(cboMotivoSunat.BoundText) + "|" + Format(RsPrinter!fEmision, "dd/mm/yyyy") + "|" + RsPrinter!tDoc + "|" + _
                                 RsPrinter!Serie + "|" + RsPrinter!corre + "||" + RsPrinter!Doc + "|" _
                                 ; IIf(IsNull(RsPrinter!NumeroR), "", Trim(RsPrinter!NumeroR)) + "|" + IIf(IsNull(RsPrinter!RazonSocial), "", Trim(RsPrinter!RazonSocial)) + "|" + _
                                 RsPrinter!BaseImOpGra + "|" + RsPrinter!Exonerada + "|" + _
                                 RsPrinter!Inafecta + "|" + RsPrinter!Isc + "|" + _
                                 RsPrinter!Igv + "|" + RsPrinter!OtrosTrib + "|" + _
                                 RsPrinter!ImporteT + "|" + RsPrinter!tipodocumentomodifica + "|" + _
                                 RsPrinter!seriemodifica + "|" + RsPrinter!numeromodifica
                    RsPrinter.MoveNext
                  Loop
                  Close #1
                  'RsSunat.ExportOptions.DiskFileName = cmdSave.FileName
                  'RsSunat.Export False
               
          Case Is = 3 ' Salir
               Unload Me
   End Select
   cmdOpcion(Index).Enabled = True
Cancelar:
   Exit Sub
End Sub

Private Sub chkCliente_Click()
   If chkCliente.Value = 1 Then
      sCliente = ""
      txtCliente.Text = ""
      cmdBusca.Enabled = False
   Else
      cmdBusca.Enabled = True
   End If
End Sub

Private Sub chkEstado_Click()
   If chkEstado.Value = 1 Then
      cboEstado.Enabled = False
      cboEstado.Text = ""
   Else
      cboEstado.Enabled = True
   End If
End Sub

Private Sub chkTipoDocumento_Click()
   If chkTipoDocumento.Value = 1 Then
      cboTipoDocumento.Enabled = False
      cboTipoDocumento.Text = ""
   Else
      cboTipoDocumento.Enabled = True
   End If
End Sub

'Private Sub chkTipoPago_Click()
'   If chkTipoPago.value = 1 Then
'      cboTipoPago.Enabled = False
'      cboTipoPago.Text = ""
'   Else
'      cboTipoPago.Enabled = True
'   End If
'End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.Value > dtpFecFin.Value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub Form_Load()
   Centrar Me
   dtpFecIni.Value = Date
   dtpFecFin.Value = Date
   cmdBusca.Enabled = False
   cboOrden.Text = cboOrden.List(2)
    
   cboTipoDocumento.Enabled = False
   cboEstado.Enabled = False
   
    
    
    LlenaCombos
  
  ' sReporte = dbTemporal(sCaja, 20, "fRegistro", "datetime", "tDocumento", "varchar(15)", "tCodigoCliente", "nVarchar(5)", _
                        "nNeto", "float", "nImpuesto1", "float", "nImpuesto2", "float", "nImpuesto3", "float", "nVenta", "float", "nRecargo", "float", _
                        "nDescuento", "float", "tEstadoDocumento", "varchar(2)", "tTipoDocumento", "varchar(2)", "tTemporal", "varchar(15)", _
                        "nNetoSuma", "float", "nImpuestoSuma1", "float", "nImpuestoSuma2", "float", "nImpuestoSuma3", "float", "nVentaSuma", "float", "nRecargoSuma", "float", "nDescuentoSuma", "float")
End Sub



Public Sub Genera1()
    '--------CESAR-------Registro Sunat
    Dim oComando As clsComando
    
    
    fInicio = Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format("00:00", "HH:mm")
    fFinal = Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format("23:59", "HH:mm")
    
    
    Screen.MousePointer = vbHourglass
    Set oComando = New clsComando
    
   '--------Emite el reporte sunat
   ' If optOpcion(0).value = True Then
       If Not oComando.CreateCmdSp("spRep_Sunattxt", Cn) Then
         Set oComando = Nothing
         Exit Sub
       End If
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 20, sCliente
    oComando.CreateParameter "@tTipoDoc", adVarChar, adParamInput, 20, cboTipoDocumento.BoundText
    oComando.CreateParameter "@tEstadoDoc", adVarChar, adParamInput, 20, cboEstado.BoundText
    oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 5, cboCaja.BoundText
    oComando.CreateParameter "@sOrden", adVarChar, adParamInput, 125, cboOrden.Text
    'oComando.CreateParameter "@flagRegVenta", adBoolean, adParamInput, 1, Me.chkRegistroVenta.value
    'oComando.CreateParameter "@flagRedondeo", adBoolean, adParamInput, 1, IIf(UCase(cboRedondeo.Text) = "DOS DECIMALES", True, False)
    'oComando.CreateParameter "@diaContable", adBoolean, adParamInput, 1, Me.chkDiaContable.value
            
   'End If
    
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    Set RsPrinter = oComando.GetSP()
    
     'RsSunat.DiscardSavedData
     'RsSunat.Database.SetDataSource RsPrinter
     'RsSunat.ReportTitle = IIf(chkDiaContable.value = 1, "Por Dia Contable", "") & " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy")
     'RsSunat.Text21.SetText sRazonSocial
     'RsSunat.Text22.SetText sRUC
     'RsSunat.Text23.SetText localConectado

     'frmEmite.CRViewer.ReportSource = RsSunat
    
     Screen.MousePointer = vbDefault
    
    '/////////////////////////////////
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Set RsPrinter = Nothing
   Set frmRepRegistroVenta = Nothing
   Set RsCliente = Nothing
   Set RsCaja = Nothing
   Set RsEstado = Nothing
   Set RsTipoDocumento = Nothing
   Set RsTipoPago = Nothing
End Sub
'DiaContable



