VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRepAsistencia 
   Caption         =   "Reporte de Asistencia"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
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
      Left            =   720
      Picture         =   "frmRepAsistencia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
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
      Left            =   2175
      Picture         =   "frmRepAsistencia.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
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
      Left            =   3600
      Picture         =   "frmRepAsistencia.frx":0A64
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
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
      Index           =   4
      Left            =   5160
      Picture         =   "frmRepAsistencia.frx":0B56
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones para Consulta"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.OptionButton RbRptDia 
         Caption         =   "Día"
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   2880
         Width           =   975
      End
      Begin VB.OptionButton RbRptUsuario 
         Caption         =   "Usuario"
         Height          =   255
         Left            =   1200
         TabIndex        =   18
         Top             =   2880
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Rango de fecha"
         Height          =   1215
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   4935
         Begin MSComCtl2.DTPicker DTPicker2 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "hh:mm:ss AMPM"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   16
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:mm 'HRS'"
            Format          =   57147395
            UpDown          =   -1  'True
            CurrentDate     =   43896.9999884259
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   3480
            TabIndex        =   15
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:mm 'HRS'"
            Format          =   57147395
            UpDown          =   -1  'True
            CurrentDate     =   43896
         End
         Begin MSComCtl2.DTPicker dtpFecFin 
            Height          =   375
            Left            =   1410
            TabIndex        =   10
            Top             =   720
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
            Format          =   57147393
            CurrentDate     =   43896.9993055556
         End
         Begin MSComCtl2.DTPicker dtpFecIni 
            Height          =   375
            Left            =   1410
            TabIndex        =   11
            Top             =   240
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
            Format          =   57147393
            CurrentDate     =   43896.2083333333
         End
         Begin VB.Label LblCodigo 
            Height          =   375
            Left            =   3480
            TabIndex        =   14
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
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
            Left            =   240
            TabIndex        =   13
            Top             =   810
            Width           =   1125
         End
         Begin VB.Label Label4 
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
            TabIndex        =   12
            Top             =   360
            Width           =   1230
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Buscar"
         Height          =   1215
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5535
         Begin VB.OptionButton RbTodos 
            Caption         =   "Todos"
            Height          =   255
            Left            =   4400
            TabIndex        =   22
            Top             =   810
            Width           =   975
         End
         Begin VB.OptionButton RbUsuario 
            Caption         =   "Usuario"
            Height          =   255
            Left            =   4400
            TabIndex        =   21
            Top             =   520
            Width           =   1095
         End
         Begin VB.OptionButton RbMesero 
            Caption         =   "Mesero"
            Height          =   255
            Left            =   4400
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   975
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
            Left            =   3550
            Picture         =   "frmRepAsistencia.frx":1378
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   465
            Width           =   765
         End
         Begin VB.TextBox txtNom 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   3375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   7
            Top             =   525
            Width           =   855
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Ordenar por:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2880
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmRepAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsMarca As Recordset
Dim rsReporte As Recordset
Dim Detallado As New dsrAsistencia
Dim Detalladoxfecha_dia As New dsrAsistencia2

Private Sub CheckxFecha_Click()
If CheckxFecha Then
    dtpFecIni.Enabled = True
    dtpFecFin.Enabled = True
Else
    dtpFecIni.Enabled = False
    dtpFecFin.Enabled = False
End If
End Sub

Private Sub cmdBusca_Click(Index As Integer)
     If RbMesero.Value = True Then
       Isql = "select codigo, Descripcion from vMozo "
     Else
        Isql = " select tCodigoUsuario as codigo ,tDetallado as Descripcion from TUSUARIO"
     End If

    frmBusca.cboCriterio.Enabled = True
    frmBusca.nPredeterm = 1
    Call ConfGrilla(2, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                           "Usuario", 2, "Descripcion", 6000, 0, 0, "")
    frmBusca.Show vbModal
    If Not wEnter Then
       Exit Sub
    End If
    LblCodigo.Caption = sCodigo
    txtNom.Text = sDescrip
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
 
 Select Case Index
        Case Is = 0 ' Preview
        Genera
            
            If rsReporte.EOF = True Then
               Screen.MousePointer = vbDefault
               MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
               Exit Sub
            End If
            
           frmEmite.CRViewer.DisplayGroupTree = True
           
           If RbRptUsuario.Value = True Then
                Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                Detallado.PaperOrientation = crPortrait
            Else
                Detalladoxfecha_dia.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                Detalladoxfecha_dia.PaperOrientation = crPortrait
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
            
             If RbRptUsuario.Value = True Then
                Detallado.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                Detallado.PaperOrientation = crPortrait
                Detallado.PrintOut
            Else
                Detalladoxfecha_dia.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                Detalladoxfecha_dia.PaperOrientation = crPortrait
                Detalladoxfecha_dia.PrintOut
            End If
                
        Case Is = 2 ' Salir
          
            Unload Me
          
'        Case Is = 3 ' Exportar
'            Genera
'            Screen.MousePointer = vbDefault
'            If rsReporte.EOF = True Then
'               MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
'               Exit Sub
'            End If
'            If optOpcion(0).Value Then
'                Detallado.ExportOptions.FormatType = 21
'                Detallado.ExportOptions.DestinationType = 1
'                cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
'                cmdSave.ShowSave
'                If cmdSave.FileName = "" Then
'                   Exit Sub
'                End If
'                Detallado.ExportOptions.DiskFileName = cmdSave.FileName
'                Detallado.Export False
'            Else
'                Consolidado.ExportOptions.FormatType = 21
'                Consolidado.ExportOptions.DestinationType = 1
'                cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
'                cmdSave.ShowSave
'                If cmdSave.FileName = "" Then
'                   Exit Sub
'                End If
'                Consolidado.ExportOptions.DiskFileName = cmdSave.FileName
'                Consolidado.Export False
'            End If
   End Select
     
     
   
End Sub
Public Sub Genera()
    Dim oComando As clsComando
    Dim sPrecio As String
    Screen.MousePointer = vbHourglass
    Set oComando = New clsComando
    
    
'    Set RsMarca = Lib.OpenRecordset(" spRep_Asistencia '" & fe1 & "','" & fe2 & "','" & tipousuario & "','" & LblCodigo.Caption & "'", Cn)
'
    If Not oComando.CreateCmdSp("spRep_Asistencia", Cn) Then
       Set oComando = Nothing
       Exit Sub
    End If
    
    Dim fe1, fe2 As Date
    Dim tipousuario, Usuario, hi, hf As String

   
       fe1 = dtpFecIni.Value
       fe2 = dtpFecFin.Value
    hi = DTPicker1.Value
    hf = DTPicker2.Value

    If LblCodigo.Caption = "" Then
        Usuario = ""
    Else
        Usuario = LblCodigo.Caption
    End If
    
    If RbMesero.Value = True Then
        tipousuario = "MO"
    End If
    
     If RbUsuario.Value = True Then
        tipousuario = "US"
    End If
    
    If RbTodos.Value = True Then
        tipousuario = ""
    End If
    
'    fInicio = Format(dtpFecIni.Value, "yyyy/mm/dd") & " " & Format(dtpHorIni.Value, "HH:mm")
'    fFinal = Format(dtpFecFin.Value, "yyyy/mm/dd") & " " & Format(dtpHorFin.Value, "HH:mm")

    oComando.CreateParameter "@fecha1", adDBDate, adParamInput, 10, fe1
    oComando.CreateParameter "@fecha2", adDBDate, adParamInput, 10, fe2
    oComando.CreateParameter "@ttipousuario", adVarChar, adParamInput, 2, tipousuario
    oComando.CreateParameter "@CODIGO", adVarChar, adParamInput, 100, Usuario
    
    oComando.CreateParameter "@hi", adDBTime, adParamInput, 7, hi
    oComando.CreateParameter "@hf", adDBTime, adParamInput, 7, hf
     
'
    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       Exit Sub
    End If
    
    Set rsReporte = oComando.GetSP()
  If RbRptUsuario.Value = True Then
        Detallado.DiscardSavedData
        Detallado.Database.SetDataSource rsReporte
        Detallado.ReportTitle = sTitulo
        Detallado.Text10.SetText "Reporte de Asistencia"
        Detallado.Text7.SetText localConectado
        Detallado.Text13.SetText sRazonSocial
        frmEmite.CRViewer.ReportSource = Detallado
    Else
        Detalladoxfecha_dia.DiscardSavedData
        Detalladoxfecha_dia.Database.SetDataSource rsReporte
        Detalladoxfecha_dia.ReportTitle = sTitulo
        Detalladoxfecha_dia.Text10.SetText "Reporte de Asistencia"
        Detalladoxfecha_dia.Text7.SetText localConectado
        Detalladoxfecha_dia.Text13.SetText sRazonSocial
        frmEmite.CRViewer.ReportSource = Detalladoxfecha_dia
    End If
End Sub



Private Sub Form_Load()
    dtpFecIni.Value = Date
    dtpFecFin.Value = Date
End Sub

Private Sub RbMesero_Click()
    cmdBusca(0).Enabled = True
    txtNom.Enabled = True
End Sub

Private Sub RbTodos_Click()
    cmdBusca(0).Enabled = False
     txtNom.Enabled = False
     LblCodigo.Caption = ""
    txtNom.Text = ""
End Sub

Private Sub RbUsuario_Click()
    cmdBusca(0).Enabled = True
     txtNom.Enabled = True
End Sub
