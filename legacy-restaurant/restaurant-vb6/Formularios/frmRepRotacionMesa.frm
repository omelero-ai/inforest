VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepRotacionMesa 
   Caption         =   "Rotación por Mesa"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6660
   Icon            =   "frmRepRotacionMesa.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6660
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
      Left            =   1290
      Picture         =   "frmRepRotacionMesa.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3270
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
      Left            =   2550
      Picture         =   "frmRepRotacionMesa.frx":053E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3270
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
      Left            =   5235
      Picture         =   "frmRepRotacionMesa.frx":0A70
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3270
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
      Height          =   3240
      Left            =   30
      TabIndex        =   6
      Top             =   -30
      Width           =   6570
      Begin VB.Frame Frame2 
         Caption         =   "Ordenado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1335
         TabIndex        =   20
         Top             =   2400
         Width           =   4905
         Begin VB.OptionButton optOpcion 
            Caption         =   "Por Rotación"
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
            Left            =   2700
            TabIndex        =   22
            Top             =   330
            Width           =   1545
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Por Mesa"
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
            Left            =   750
            TabIndex        =   21
            Top             =   330
            Value           =   -1  'True
            Width           =   1170
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
         Left            =   4590
         TabIndex        =   17
         Top             =   1125
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
         Left            =   4590
         TabIndex        =   12
         Top             =   720
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
         Left            =   4590
         TabIndex        =   11
         Top             =   300
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1320
         TabIndex        =   2
         Top             =   1965
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
         Left            =   1320
         TabIndex        =   0
         Top             =   1515
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
         Left            =   4605
         TabIndex        =   1
         Top             =   1500
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
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   4605
         TabIndex        =   3
         Top             =   1935
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
         Format          =   63963139
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo cboLocal 
         Height          =   315
         Left            =   1320
         TabIndex        =   13
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
         Left            =   1320
         TabIndex        =   14
         Top             =   705
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
         Left            =   1320
         TabIndex        =   18
         Top             =   1110
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
         Left            =   660
         TabIndex        =   19
         Top             =   1170
         Width           =   585
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
         Left            =   630
         TabIndex        =   16
         Top             =   765
         Width           =   615
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
         Left            =   645
         TabIndex        =   15
         Top             =   345
         Width           =   600
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
         Left            =   165
         TabIndex        =   10
         Top             =   1590
         Width           =   1080
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
         Left            =   255
         TabIndex        =   9
         Top             =   1995
         Width           =   990
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
      Left            =   3885
      Picture         =   "frmRepRotacionMesa.frx":0B62
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3270
      Visible         =   0   'False
      Width           =   1350
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   645
      Top             =   3390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepRotacionMesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNombre     As String
Dim rsReporte   As Recordset
Dim RsLocal     As Recordset
Dim RsSalon     As Recordset
Dim RsMesa      As Recordset
Dim Reporte     As New dsrRotacionMesa
Dim sCriterio   As String
Dim sOrden      As String
Dim sTitulo     As String
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
        sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & "Local ='" & cboLocal.Text & "'"
    End If
    
    If chkSalon.Value = 0 Then
       If cboSalon.Text = "" Then
          MsgBox "Debe escoger un Salón", vbCritical, sMensaje
          Exit Sub
       End If
        sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & "Salon ='" & cboSalon.Text & "'"
    End If
    
    If chkMesa.Value = 0 Then
       If cboMesa.Text = "" Then
          MsgBox "Debe escoger una Mesa", vbCritical, sMensaje
          Exit Sub
       End If
       sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & "Mesa ='" & cboMesa.Text & "'"
    End If
      
    Select Case Index
        Case Is = 0 ' Preview
            Genera
            If rsReporte.EOF = True Then
               Screen.MousePointer = vbDefault
               MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
               Exit Sub
            End If
            frmEmite.CRViewer.DisplayGroupTree = True
            Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
            Reporte.PaperOrientation = crPortrait
            frmEmite.CRViewer.ViewReport
            frmEmite.Show vbModal
               
        Case Is = 1 ' Imprimir
            Genera
            Screen.MousePointer = vbDefault
            If rsReporte.EOF = True Then
                MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                Exit Sub
            End If
            Reporte.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
            Reporte.PaperOrientation = crPortrait
            Reporte.PrintOut
               
        Case Is = 2 ' Salir
            Unload Me
          
        Case Is = 3 ' Exportar
            Genera
            Screen.MousePointer = vbDefault
            If rsReporte.EOF = True Then
               MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
               Exit Sub
            End If
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
            
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsReporte = Nothing
End Sub

Public Sub Genera()
    Dim oComando As clsComando
    Screen.MousePointer = vbHourglass
    
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spRep_Rotacion", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    If optOpcion(0).Value Then
       sOrden = "Mesa"
    Else
       sOrden = "Local, Salon, nVeces, nAsociado DESC"
    End If
    
    oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
    oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal
    oComando.CreateParameter "@sCriterio", adVarChar, adParamInput, 400, IIf(sCriterio = "", sCriterio, " Where " + sCriterio) + " Order By " + sOrden
                      
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If

    
    Set rsReporte = oComando.GetSP()
    'rsReporte.Filter = sCriterio
    'rsReporte.Sort = sOrden
    
    Reporte.DiscardSavedData
    Reporte.Database.SetDataSource rsReporte
    Reporte.Text1.SetText localConectado
    Reporte.Text13.SetText sRazonSocial
    Reporte.ReportTitle = sTitulo
    frmEmite.CRViewer.ReportSource = Reporte
End Sub

Private Sub cboLocal_Click(Area As Integer)
    cboSalon.Text = ""
    With cboSalon
         Isql = "Select * from vSalon " & IIf(cboLocal.BoundText = "", "", "where tLocal = '" & cboLocal.BoundText & "'") & " order by Descripcion "
         Set RsSalon = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSalon
    End With
End Sub

