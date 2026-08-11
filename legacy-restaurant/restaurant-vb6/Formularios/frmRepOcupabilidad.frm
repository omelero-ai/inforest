VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepOcupabilidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Ocupabilidad"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   Icon            =   "frmRepOcupabilidad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7110
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
      Left            =   1785
      Picture         =   "frmRepOcupabilidad.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4035
      Width           =   1260
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
      Left            =   3045
      Picture         =   "frmRepOcupabilidad.frx":053E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4035
      Width           =   1350
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
      Left            =   5730
      Picture         =   "frmRepOcupabilidad.frx":0A70
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4035
      Width           =   1350
   End
   Begin VB.Frame Frame1 
      Caption         =   " Opciones "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   7065
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   1635
         TabIndex        =   19
         Top             =   2790
         Width           =   1905
         Begin VB.OptionButton optOpcion 
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
            Height          =   300
            Index           =   0
            Left            =   135
            TabIndex        =   21
            Top             =   225
            Value           =   -1  'True
            Width           =   1290
         End
         Begin VB.OptionButton optOpcion 
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
            Height          =   300
            Index           =   1
            Left            =   135
            TabIndex        =   20
            Top             =   585
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   3690
         TabIndex        =   15
         Top             =   2790
         Width           =   3165
         Begin VB.OptionButton optOpcion 
            Caption         =   "Solo Horas Ocupadas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   180
            TabIndex        =   17
            Top             =   225
            Value           =   -1  'True
            Width           =   2355
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Todas las horas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   180
            TabIndex        =   16
            Top             =   585
            Width           =   1875
         End
      End
      Begin VB.CheckBox chkMesa 
         Caption         =   "Todas las Mesas"
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
         Left            =   4965
         TabIndex        =   13
         Top             =   1115
         Value           =   1  'Checked
         Width           =   1905
      End
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
         Left            =   4965
         TabIndex        =   10
         Top             =   715
         Value           =   1  'Checked
         Width           =   1905
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
         Left            =   4965
         TabIndex        =   9
         Top             =   315
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1635
         TabIndex        =   2
         Top             =   1920
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
         Left            =   1635
         TabIndex        =   0
         Top             =   1500
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
         Left            =   3660
         TabIndex        =   1
         Top             =   1485
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   3660
         TabIndex        =   3
         Top             =   1905
         Width           =   1095
         _ExtentX        =   1931
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
         Left            =   1635
         TabIndex        =   11
         Top             =   300
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
         Left            =   1635
         TabIndex        =   12
         Top             =   700
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
      Begin MSDataListLib.DataCombo cboMesa 
         Height          =   315
         Left            =   1635
         TabIndex        =   14
         Top             =   1100
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
      Begin MSComCtl2.DTPicker dtpHora 
         Height          =   345
         Left            =   1635
         TabIndex        =   18
         Top             =   2355
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
         CustomFormat    =   "HH"
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   555
         TabIndex        =   28
         Top             =   1995
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   465
         TabIndex        =   27
         Top             =   1560
         Width           =   1080
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
         Left            =   945
         TabIndex        =   26
         Top             =   360
         Width           =   600
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
         Left            =   930
         TabIndex        =   25
         Top             =   760
         Width           =   615
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mesa :"
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
         Left            =   960
         TabIndex        =   24
         Top             =   1160
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Reporte :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   23
         Top             =   2835
         Width           =   1410
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora Inicial :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   570
         TabIndex        =   22
         Top             =   2415
         Width           =   975
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
      Left            =   4380
      Picture         =   "frmRepOcupabilidad.frx":0B62
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4035
      Visible         =   0   'False
      Width           =   1350
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   1125
      Top             =   4095
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepOcupabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNombre As String
Dim sTemporal As String
Dim rsReporte As Recordset
Dim RsLocal As Recordset
Dim RsSalon As Recordset
Dim RsMesa As Recordset
Dim ReporteR As New dsrOcupabilidadR
Dim ReporteD As New dsrOcupabilidadD
Dim sCriterio As String
Dim sTitulo As String
Dim sDetalle As String
Dim sHora(48) As String
Dim HoraInicial(48) As Integer
Dim HoraFinal(48) As Integer
Dim sOrden(48) As Integer
Dim i As Integer
Dim nMesa As Double
Dim nAdulto As Double

Dim fInicio As Date
Dim fFinal As Date


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
    
    With cboMesa
         Isql = "Select * from tMesa order by tDetallado"
         Set RsMesa = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsMesa
         .DataField = "tDetallado"
         .ListField = "tDetallado"
         .BoundColumn = "tCodigoMesa"
    End With
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

Private Sub chkMesa_Click()
   If chkMesa.Value = 1 Then
      cboMesa.Enabled = False
      cboMesa.Text = ""
   Else
      cboMesa.Enabled = True
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
    sDetalle = ""
    If dtpFecIni.Value + dtpHorIni.Value > dtpFecFin.Value + dtpHorFin.Value Then
        MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
        Exit Sub
        dtpFecFin.SetFocus
    End If
    sTitulo = " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.Value, "HH:mm") & " Al " & Format(dtpFecFin.Value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.Value, "HH:mm")

    fInicio = Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm")
    fFinal = Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm")
      
    If chkLocal.Value = 0 Then
            If cboLocal.Text = "" Then
               MsgBox "Debe escoger un Local", vbCritical, sMensaje
               Exit Sub
            End If
            sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & " Local ='" & cboLocal.Text & "'"
            sDetalle = " Local: " + Me.cboLocal.Text
    End If
    
    If chkSalon.Value = 0 Then
       If cboSalon.Text = "" Then
          MsgBox "Debe escoger un Salón", vbCritical, sMensaje
          Exit Sub
       End If
        sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & " Salon ='" & cboSalon.Text & "'"
        sDetalle = sDetalle + " Salon: " + cboSalon.Text
    End If
    
    If chkMesa.Value = 0 Then
       If cboMesa.Text = "" Then
          MsgBox "Debe escoger una Mesa", vbCritical, sMensaje
          Exit Sub
       End If
       sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & " dMesa ='" & cboMesa.Text & "'"
       sDetalle = sDetalle + " Mesa: " + cboMesa.Text
    End If
          
    Select Case Index
        Case Is = 0 ' Preview
            Genera
            If rsReporte.EOF = True Then
               Screen.MousePointer = vbDefault
               MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
               Exit Sub
            End If
            frmEmite.CRViewer.DisplayGroupTree = False
            
            If optOpcion(0).Value Then
               ReporteR.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               ReporteR.PaperOrientation = crPortrait
            Else
               ReporteD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               ReporteD.PaperOrientation = crPortrait
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
            If optOpcion(0).Value Then
               ReporteR.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               ReporteR.PaperOrientation = crPortrait
               ReporteR.PrintOut
            Else
               ReporteD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
               ReporteD.PaperOrientation = crPortrait
               ReporteD.PrintOut
            End If
               
        Case Is = 2 ' Salir
            Unload Me
          
        Case Is = 3 ' Exportar
            Genera
            Screen.MousePointer = vbDefault
            If rsReporte.EOF = True Then
               MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
               Exit Sub
            End If
            
            If optOpcion(0).Value Then
               ReporteR.ExportOptions.FormatType = 21
               ReporteR.ExportOptions.DestinationType = 1
               cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
               cmdSave.ShowSave
               If cmdSave.FileName = "" Then
                  Exit Sub
               End If
               ReporteR.ExportOptions.DiskFileName = cmdSave.FileName
               ReporteR.Export False
            Else
               ReporteD.ExportOptions.FormatType = 21
               ReporteD.ExportOptions.DestinationType = 1
               cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
               cmdSave.ShowSave
               If cmdSave.FileName = "" Then
                  Exit Sub
               End If
               ReporteD.ExportOptions.DiskFileName = cmdSave.FileName
               ReporteD.Export False
            End If
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
    dtpHora.Value = Date
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsReporte = Nothing
End Sub

Public Sub Genera()
    Dim oComando As clsComando
    Screen.MousePointer = vbHourglass
    
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_Ocupabilidad", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@nHora", adInteger, adParamInput, 10, Int(Hour(dtpHora))
    oComando.CreateParameter "@sCriterio", adVarChar, adParamInput, 400, sCriterio
    oComando.CreateParameter "@Opcion", adInteger, adParamInput, 10, IIf(optOpcion(0), 0, 1)
                      
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
        
    Set rsReporte = oComando.GetSP()
    'rsReporte.Filter = sCriterio
        
    If optOpcion(0).Value Then
       ReporteR.DiscardSavedData
       ReporteR.Database.SetDataSource rsReporte
       ReporteR.ReportTitle = sTitulo
       ReporteR.Text5.SetText sRazonSocial
       ReporteR.Text17.SetText localConectado
       ReporteR.Text18.SetText sDetalle
       frmEmite.CRViewer.ReportSource = ReporteR
    Else
       ReporteD.DiscardSavedData
       ReporteD.Database.SetDataSource rsReporte
       ReporteD.ReportTitle = sTitulo
       ReporteD.Text15.SetText sRazonSocial
       ReporteD.Text18.SetText sDetalle
       ReporteD.Text1.SetText localConectado
       frmEmite.CRViewer.ReportSource = ReporteD
    End If
End Sub

Private Sub cboLocal_Click(Area As Integer)
    cboSalon.Text = ""
    With cboSalon
         Isql = "Select * from vSalon " & IIf(cboLocal.BoundText = "", "", "where tLocal = '" & cboLocal.BoundText & "'") & " order by Descripcion "
         Set RsSalon = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSalon
    End With
End Sub
