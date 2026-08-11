VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRepRegistroComprasIntegrado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Compras - Integrado"
   ClientHeight    =   6015
   ClientLeft      =   1680
   ClientTop       =   3405
   ClientWidth     =   10905
   Icon            =   "frmRepRegistroComprasIntegrado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
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
      Left            =   7635
      Picture         =   "frmRepRegistroComprasIntegrado.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5340
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
      Left            =   6195
      Picture         =   "frmRepRegistroComprasIntegrado.frx":023C
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5340
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
      Left            =   4755
      Picture         =   "frmRepRegistroComprasIntegrado.frx":0304
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5340
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
      Picture         =   "frmRepRegistroComprasIntegrado.frx":0B26
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5340
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
      Height          =   5235
      Left            =   30
      TabIndex        =   19
      Top             =   30
      Width           =   10845
      Begin VB.ComboBox CboCondicion 
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "frmRepRegistroComprasIntegrado.frx":1058
         Left            =   1530
         List            =   "frmRepRegistroComprasIntegrado.frx":105A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1230
         Width           =   2970
      End
      Begin VB.CheckBox chkCondicion 
         Caption         =   "Todos "
         Height          =   195
         Left            =   4545
         TabIndex        =   5
         Top             =   1305
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   7170
         Top             =   4725
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "Xls"
      End
      Begin VB.Frame Frame5 
         Caption         =   "Formato de Reporte"
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   90
         TabIndex        =   27
         Top             =   1650
         Width           =   6090
         Begin VB.OptionButton optFormato 
            Caption         =   "Registro de compras"
            Height          =   195
            Index           =   2
            Left            =   4230
            TabIndex        =   8
            Top             =   390
            Width           =   1755
         End
         Begin VB.OptionButton optFormato 
            Caption         =   "Por familia / subfamilia"
            Height          =   195
            Index           =   1
            Left            =   2175
            TabIndex        =   7
            Top             =   390
            Width           =   1920
         End
         Begin VB.OptionButton optFormato 
            Caption         =   "Por tipo de documento"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   390
            Value           =   -1  'True
            Width           =   2010
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Filtrado Por: "
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   90
         TabIndex        =   26
         Top             =   2655
         Width           =   6090
         Begin VB.OptionButton OptTipoFecha 
            Caption         =   "Emisión"
            Height          =   195
            Index           =   1
            Left            =   2190
            TabIndex        =   10
            Top             =   405
            Width           =   1290
         End
         Begin VB.OptionButton OptTipoFecha 
            Caption         =   "Ingreso"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   9
            Top             =   420
            Value           =   -1  'True
            Width           =   1635
         End
         Begin VB.OptionButton OptTipoFecha 
            Caption         =   "Procesamiento"
            Height          =   195
            Index           =   2
            Left            =   4245
            TabIndex        =   11
            Top             =   405
            Width           =   1635
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Tipo de Reporte "
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   75
         TabIndex        =   23
         Top             =   3705
         Width           =   6090
         Begin VB.OptionButton optOpcion 
            Caption         =   "Por sector"
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   12
            Top             =   420
            Width           =   1500
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Por local"
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   1
            Left            =   2190
            TabIndex        =   13
            Top             =   420
            Value           =   -1  'True
            Width           =   2475
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Servidores en línea"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5160
         Left            =   6225
         TabIndex        =   20
         Top             =   0
         Width           =   4545
         Begin VB.CheckBox chkTodo 
            Caption         =   "Seleccionar todo"
            Height          =   300
            Left            =   2880
            TabIndex        =   21
            Top             =   4830
            Width           =   1575
         End
         Begin MSComctlLib.ListView lstServidores 
            Height          =   4605
            Left            =   90
            TabIndex        =   22
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   8123
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
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   330
         Left            =   1530
         TabIndex        =   2
         Top             =   780
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Format          =   60751873
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   330
         Left            =   1530
         TabIndex        =   0
         Top             =   330
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Format          =   60751873
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   330
         Left            =   3150
         TabIndex        =   1
         Top             =   330
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
         Format          =   60751875
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   330
         Left            =   3150
         TabIndex        =   3
         Top             =   780
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
         Format          =   60751875
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Condición :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   675
         TabIndex        =   28
         Top             =   1275
         Width           =   795
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   25
         Top             =   390
         Width           =   990
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   555
         TabIndex        =   24
         Top             =   840
         Width           =   915
      End
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
      Left            =   1875
      Picture         =   "frmRepRegistroComprasIntegrado.frx":105C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5340
      Width           =   1455
   End
End
Attribute VB_Name = "frmRepRegistroComprasIntegrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsEmpresa     As Recordset
Dim RsDocumento   As Recordset
Dim RsTingreso    As Recordset
Dim RsDatos       As Recordset
Dim rsReporte     As Recordset
Dim sCriterio     As String
Dim sCriterioN    As String
Dim sNombre       As String
Dim sTitulo       As String
Dim sTituloG      As String
Dim strTipoFiltro As String
Dim CrtReporteRC  As New CrtRegComprasIntegrado
Dim CrtReporteTD  As New CrtComprasxTipoDocumentoIntegrado
Dim CrtReporteTP  As New CrtComprasxTipoProductoIntegrado

Dim RsFamilia     As Recordset
Dim RsSubFamilia  As Recordset
Dim RsEstado      As Recordset

Sub LlenaCombo()
    With CboCondicion
        .AddItem "Pendientes de Pago"       'nTotal = nSaldo
        .AddItem "Amortizados"              'nTotal > nSaldo
        .AddItem "Cancelados"               'nSaldo = 0
    End With
End Sub

Sub CreaTemporal()
    If optFormato(2).value Then
        sNombre = dbTemporal(sUsuario, 39, "tLocalidad", "nVarChar(100)", _
                                           "ncorrela", "nVarChar(30)", _
                                           "FEmision", "DateTime", "fProgramacion", "DateTime", _
                                           "tTipoDocumento", "nVarChar(2)", "tDocumento", "nVarChar(15)", _
                                           "tGlosa", "nVarChar(50)", _
                                           "tIndicadorTributario", "nVarChar(18)", _
                                           "tRazonSocial", "nVarChar(100)", _
                                           "tMoneda", "nVarChar(2)", _
                                           "Inafecta", "Float", _
                                           "Afecta", "Float", _
                                           "nImpuesto1", "Float", "nImpuesto2", "Float", "nImpuesto3", "Float", _
                                           "nTotal", "Float", "nPercepcion", "Float", "nGranTotal", "Float", _
                                           "tEstadoDocumento", "nVarChar(2)", _
                                           "tCodigoProveedor", "nVarChar(5)", _
                                           "tTemporal", "nvarchar(15)", _
                                           "nInafectaSuma", "float", "nAfectaSuma", "float", _
                                           "nImpuestoSuma1", "float", "nImpuestoSuma2", "float", "nImpuestoSuma3", "float", _
                                           "nTotalSuma", "float", _
                                           "nPercepcionSuma", "float", _
                                           "nGranTotalSuma", "float", _
                                           "fRegistro", "datetime", "fProcesa", "datetime", _
                                           "nCambio", "Float", _
                                           "fDocRef", "DATETIME", _
                                           "tTipoDocRef", "NVARCHAR(2)", "tSerieDocRef", "NVARCHAR(5)", "tNumDocRef", "NVARCHAR(9)", _
                                           "tTipoIngreso", "NVARCHAR(4)", "nNeto", "FLOAT", "fIngreso", "DateTime")
                                           
'                                           "nCambio", "Float", _

    ElseIf optFormato(0).value Then
        sNombre = dbTemporal(sUsuario, 10, "tLocalidad", "nVarChar(100)", _
                                           "tDocumento", "nVarChar(200)", _
                                           "Inafecta", "Float", _
                                           "Afecta", "Float", _
                                           "nImpuesto1", "Float", _
                                           "nImpuesto2", "Float", _
                                           "nImpuesto3", "Float", _
                                           "nTotal", "Float", _
                                           "Desde", "nVarchar(100)", _
                                           "Hasta", "nVarchar(100)")
    Else
        sNombre = dbTemporal(sUsuario, 9, "tLocalidad", "nVarChar(100)", _
                                          "tFamilia", "nVarChar(100)", _
                                          "tSubFamilia", "nVarChar(100)", _
                                          "Inafecta", "Float", _
                                          "Afecta", "Float", _
                                          "nImpuesto1", "Float", _
                                          "nImpuesto2", "Float", _
                                          "nImpuesto3", "Float", _
                                          "nTotal", "Float")
    End If
End Sub

Private Sub configuraListaServidores()
    Dim rsS As New ADODB.Recordset
    Dim Item As ListItem
    Dim X As Integer
    
    Set rsS = devuelveServidoresConectados
    
    With lstServidores
        .Checkboxes = True
        .ColumnHeaders.Clear
        .ListItems.Clear
    
        .ColumnHeaders.Add , , " Código ", 800
        .ColumnHeaders.Add , , " Sector ", 1100
        .ColumnHeaders.Add , , " Local ", 1100
        .ColumnHeaders.Add , , " Estado ", 1100
        .ColumnHeaders.Add , , "Ip", 0
        .ColumnHeaders.Add , , "Bd", 0
        .View = lvwReport
        .GridLines = True
     
        rsS.MoveFirst
        For X = 0 To rsS.RecordCount - 1
            Set Item = .ListItems.Add(, , rsS.Fields(0))
            Item.SubItems(1) = rsS.Fields(1)
            Item.SubItems(2) = rsS.Fields(2)
            Item.SubItems(3) = rsS.Fields(5)
            Item.SubItems(4) = rsS.Fields(3)
            Item.SubItems(5) = rsS.Fields(6)
            rsS.MoveNext
        Next X
    End With
End Sub

Public Sub GeneraDetalle()
    Dim oComando As clsComando
    Dim w, cl
    Dim conectado As Boolean
    Dim sector As String, llocal As String, Localidad As String, sOrden As String
    
    Set oComando = New clsComando

    If sNombre <> "" Then Cn.Execute "Delete From " & sNombre
        
    For w = 1 To lstServidores.ListItems.Count
        If lstServidores.ListItems.Item(w).Checked = True Then
            conectado = validaConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))
            If conectado = False Then
                MsgBox "Error Fatal: No se pudo conectar con  " & lstServidores.ListItems.Item(w).SubItems(2), vbCritical, sMensaje
                lstServidores.ListItems.Item(w).Checked = False
                GoTo sigue
            End If
            
            If optFormato(0).value Then 'Por Tipo de Documento
                If Not oComando.CreateCmdSp("spRep_RegCompraTipoDocumentoIntegrado", devuelveConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))) Then
                    Set oComando = Nothing
                    Exit Sub
                End If
            ElseIf optFormato(1).value Then 'Por Tipo de Producto
                If Not oComando.CreateCmdSp("spRep_RegCompraTipoProductoIntegrado", devuelveConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))) Then
                    Set oComando = Nothing
                    Exit Sub
                End If
            Else 'Registro de Compras
                If Not oComando.CreateCmdSp("spRep_RegCompraIntegrado", devuelveConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))) Then
                    Set oComando = Nothing
                    Exit Sub
                End If
            End If
            
            oComando.CreateParameter "@Local", adVarChar, adParamInput, 2, lstServidores.ListItems.Item(w)
            oComando.CreateParameter "@TipoFecha", adVarChar, adParamInput, 1, IIf(OptTipoFecha(0).value, "I", IIf(OptTipoFecha(1).value, "E", "P"))
            oComando.CreateParameter "@FechaIni", adVarChar, adParamInput, 20, Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "HH:mm")
            oComando.CreateParameter "@FechaFin", adVarChar, adParamInput, 20, Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(dtpHorFin.value, "HH:mm")
            oComando.CreateParameter "@Estado", adVarChar, adParamInput, 2, IIf(CboCondicion.ListIndex = -1, "*", Choose(CboCondicion.ListIndex + 1, "P", "A", "C"))
            If Not oComando.GetParamOK Then
                Set oComando = Nothing
                Exit Sub
            End If
            cl = lstServidores.ListItems.Item(w)
            sector = lstServidores.ListItems.Item(w).SubItems(1)
            llocal = lstServidores.ListItems.Item(w).SubItems(2)
            Set RsDatos = oComando.GetSP()
            With RsDatos
                While Not RsDatos.EOF
                    If optOpcion(0).value Then
                        Localidad = UCase(sector)
                    Else
                        Localidad = UCase(llocal) '& " (" & sector & ")"
                    End If
                    
                    If optFormato(0).value Then 'Por Tipo de Documento
                        Isql = "INSERT INTO " & sNombre & "(tLocalidad, tDocumento, Inafecta, Afecta, nImpuesto1, nImpuesto2, nImpuesto3, nTotal, Desde, Hasta) " & _
                               "VALUES('" & Localidad & "','" & !tDocumento & "'," & !Inafecta & "," & !Afecta & ", " & !nImpuesto1 & "," & !nImpuesto2 & "," & !nImpuesto3 & "," & !nTotal & ",'" & !Minimo & "','" & !Maximo & "')"
                    ElseIf optFormato(1).value Then 'Por Tipo de Producto
                        Isql = "INSERT INTO " & sNombre & "(tLocalidad, tFamilia, tSubFamilia, Inafecta, Afecta, nImpuesto1, nImpuesto2, nImpuesto3, nTotal) " & _
                               "VALUES('" & Localidad & "','" & !tFamilia & "','" & !tsubFamilia & "'," & !Inafecta & "," & !Afecta & ", " & !nImpuesto1 & "," & !nImpuesto2 & "," & !nImpuesto3 & "," & !nTotal & ")"
                    Else
                        'Isql = "INSERT INTO " & sNombre & "(tLocalidad,nCorrela,fEmision,fProgramacion,tTipoDocumento,tDocumento,tGlosa,tIndicadorTributario,tRazonSocial,tMoneda,Inafecta,Afecta,nImpuesto1,nImpuesto2,nImpuesto3,nTotal,nPercepcion,nGranTotal,tEstadoDocumento,tCodigoProveedor,tTemporal,nInafectaSuma,nAfectaSuma,nImpuestoSuma1,nImpuestoSuma2,nImpuestoSuma3,nTotalSuma,nPercepcionSuma,nGranTotalSuma,fRegistro,fProcesa,nCambio,fDocRef,tTipoDocRef,tSerieDocRef,tNumDocRef,tTipoIngreso,nNeto,fIngreso) " & _
                        '       "VALUES('" & Localidad & "','" & !nCorrela & "', Convert(Datetime, '" & !fEmision & "', 103),'" & !tDocumento & "','" & !tTipoDocumento & "','" & !tGlosa & "','" & !tRazonSocial & "','" & !tIndicadorTributario & "','" & !tMoneda & "'," & !nCambio & "," & !Inafecta & "," & !Afecta & ", " & !nImpuesto1 & "," & !nImpuesto2 & "," & !nImpuesto3 & "," & !nTotal & "," & !tEstadoDocumento & "," & !tCodigoProveedor & ",'" & !tTemporal & "'," & !nInafectaSuma & "," & !nAfectaSuma & "," & !nImpuestoSuma1 & "," & !nImpuestoSuma2 & "," & !nImpuestoSuma3 & "," & !nTotalSuma & ", Convert(Datetime, '" & Format(!fRegistro, "dd/MM/yyyy hh:mm:ss am/pm") & "', 103), Convert(Datetime, '" & Format(!fProcesa, "dd/MM/yyyy hh:mm:ss am/pm") & "',103))"
                               
                        Isql = "INSERT INTO " & sNombre & "(tLocalidad,nCorrela,fEmision,fProgramacion,tTipoDocumento,tDocumento,tGlosa,tIndicadorTributario,tRazonSocial,tMoneda,Inafecta,Afecta,nImpuesto1,nImpuesto2,nImpuesto3,nTotal,nPercepcion,nGranTotal,tEstadoDocumento,tCodigoProveedor,tTemporal,nInafectaSuma,nAfectaSuma,nImpuestoSuma1,nImpuestoSuma2,nImpuestoSuma3,nTotalSuma,nPercepcionSuma,nGranTotalSuma,fRegistro,fProcesa,nCambio,fDocRef,tTipoDocRef,tSerieDocRef,tNumDocRef,tTipoIngreso,nNeto,fIngreso) " & _
                               " VALUES('" & Localidad & "','" & !nCorrela & "',Convert(Datetime,'" & !fEmision & "', 103),Convert(Datetime,'" & !fProgramacion & "', 103),'" & !tTipoDocumento & "','" & !tDocumento & "','" & !tGlosa & "','" & !tIndicadorTributario & "','" & !tRazonSocial & "','" & !tMoneda & "'," & !Inafecta & "," & !Afecta & "," & !nImpuesto1 & "," & !nImpuesto2 & "," & !nImpuesto3 & "," & !nTotal & "," & !nPercepcion & "," & !nGranTotal & "," & !tEstadoDocumento & "," & !tCodigoProveedor & ",'" & !tTemporal & "'," & !nInafectaSuma & "," & !nAfectaSuma & "," & !nImpuestoSuma1 & "," & !nImpuestoSuma2 & "," & !nImpuestoSuma3 & "," & !nTotalSuma & "," & !nPercepcionSuma & "," & !nGranTotalSuma & ", " & _
                               " Convert(Datetime, '" & Format(!fRegistro, "dd/MM/yyyy hh:mm:ss am/pm") & "', 103),Convert(Datetime, '" & Format(!fProcesa, "dd/MM/yyyy hh:mm:ss am/pm") & "',103)," & !nCambio & ",case when Convert(Datetime,'" & Format(!fDocRef, "dd/MM/yyyy hh:mm:ss am/pm") & "', 103)='01/01/1900' then NULL else Convert(Datetime,'" & Format(!fDocRef, "dd/MM/yyyy hh:mm:ss am/pm") & "', 103) end ,'" & !tTipoDocRef & "','" & !tSerieDocRef & "','" & !tNumDocRef & "','" & !tTipoIngreso & "'," & !nNeto & ",Convert(Datetime, '" & Format(!fIngreso, "dd/MM/yyyy hh:mm:ss am/pm") & "', 103))"
                               '" & !nCorrela & "',Convert(Datetime,'" & !fEmision & "', 103),Convert(Datetime,'" & !fProgramacion & "', 103),'" & !tTipoDocumento & "','" & !tDocumento & "','" & !tGlosa & "','" & !tIndicadorTributario & "','" & !tRazonSocial & "','" & !tMoneda & "'," & !Inafecta & "," & !Afecta & "," & !nImpuesto1 & "," & !nImpuesto2 & "," & !nImpuesto3 & "," & !nTotal & "," & !nPercepcion & "," & !nGranTotal & "," & !tEstadoDocumento & "," & !tCodigoProveedor & ",'" & !tTemporal & "'," & !nInafectaSuma & "," & !nAfectaSuma & "," & !nImpuestoSuma1 & "," & !nImpuestoSuma2 & "," & !nImpuestoSuma3 & "," & !nTotalSuma & "," & !nPercepcionSuma & "," & !nGranTotalSuma & ", _
                               Convert(Datetime, '" & Format(!fRegistro, "dd/MM/yyyy hh:mm:ss am/pm") & "', 103),Convert(Datetime, '" & Format(!fProcesa, "dd/MM/yyyy hh:mm:ss am/pm") & "',103)," & !nCambio & ",Convert(Datetime, '" & Format(!fDocRef, "dd/MM/yyyy hh:mm:ss am/pm") & "', 103),'" & !tTipoDocRef & "','" & !tSerieDocRef & "','" & !tNumDocRef & "','" & !tTipoIngreso & "'," & !nNeto & ",'" & !fIngreso

                               
                    End If
                    Debug.Print Isql
                    Cn.Execute Isql
                    RsDatos.MoveNext
                Wend
            End With
        End If
sigue:
    Next w

    If optFormato(2).value Then
        sOrden = " Order By ncorrela"
        Isql = "SELECT tLocalidad AS Grupo, * FROM " & sNombre & sOrden
    ElseIf optFormato(1).value Then
        sOrden = " Order By tFamilia, tSubFamilia"
        Isql = "SELECT tLocalidad AS Grupo, tFamilia, tSubFamilia, SUM(Inafecta) AS Inafecta, SUM(Afecta) AS Afecta, SUM(nImpuesto1) AS nImpuesto1, SUM(nImpuesto2) AS nImpuesto2, SUM(nImpuesto3) AS nImpuesto3, SUM(nTotal) AS nTotal FROM " & sNombre & " GROUP BY tLocalidad, tFamilia, tSubFamilia " & sOrden
    Else
        sOrden = " Order By tDocumento"
        If optOpcion(0).value Then 'Zonal
            Isql = "SELECT tLocalidad AS Grupo, tDocumento, SUM(Inafecta) AS Inafecta, SUM(Afecta) AS Afecta, SUM(nImpuesto1) AS nImpuesto1, SUM(nImpuesto2) AS nImpuesto2, SUM(nImpuesto3) AS nImpuesto3, SUM(nTotal) AS nTotal, MAX(Desde) AS Desde, MAX(Hasta) AS Hasta FROM " & sNombre & " GROUP BY tLocalidad, tDocumento " & sOrden
        Else
            Isql = "SELECT tLocalidad AS Grupo, tDocumento, SUM(Inafecta) AS Inafecta, SUM(Afecta) AS Afecta, SUM(nImpuesto1) AS nImpuesto1, SUM(nImpuesto2) AS nImpuesto2, SUM(nImpuesto3) AS nImpuesto3, SUM(nTotal) AS nTotal, Desde, Hasta FROM " & sNombre & " GROUP BY tLocalidad, tDocumento, Desde, Hasta " & sOrden
        End If
    End If
    Set rsReporte = Lib.OpenRecordset(Isql, Cn)
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

Private Sub obtieneListaServidores()
    Dim rsS As New ADODB.Recordset
    Dim Item As ListItem
    Dim X As Integer
    Dim xx As String
    
    Set rsS = rsListaServidores
    Me.lstServidores.Checkboxes = True
    lstServidores.ColumnHeaders.Clear
    Me.lstServidores.ListItems.Clear

    With lstServidores
        .ColumnHeaders.Add , , " Código ", 800
        .ColumnHeaders.Add , , " Sector ", 1100
        .ColumnHeaders.Add , , " Local ", 1100
        .ColumnHeaders.Add , , " Estado ", 1100
        .ColumnHeaders.Add , , "Ip", 0
        .ColumnHeaders.Add , , "Bd", 0
        .View = lvwReport
        .GridLines = True
    
        rsS.MoveFirst
        For X = 0 To rsS.RecordCount - 1
            Set Item = .ListItems.Add(, , rsS.Fields(0))
            Item.SubItems(1) = rsS.Fields(1) 'Sector
            Item.SubItems(2) = rsS.Fields(2) 'Local
            Item.SubItems(3) = rsS.Fields(5) 'Estado
            'Item.SubItems(4) = rsS.Fields(3) 'Ip
            Item.SubItems(4) = rsS.Fields(6) 'Ip Almacen
            'Item.SubItems(5) = rsS.Fields(6) 'Base de Datos
            Item.SubItems(5) = rsS.Fields(7) 'Base de Datos Almacen
        rsS.MoveNext
        Next X
    End With
End Sub

Private Sub chkCondicion_Click()
    If chkCondicion.value Then
        CboCondicion.BackColor = &H8000000F
        CboCondicion.Enabled = False
        CboCondicion.ListIndex = -1
    Else
        CboCondicion.BackColor = vbWhite
        CboCondicion.Enabled = True
    End If
End Sub

Private Sub chkTodo_Click()
    activaCheck IIf(chkTodo.value = 1, True, False)
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

Private Sub cmdOpcion_Click(Index As Integer)
Dim lSw As Boolean
Dim FecIni As String
Dim FecFin  As String
Dim w

    If Index <> 3 Then
        
        CreaTemporal
            
        If chkCondicion.value = 0 Then
            If CboCondicion.ListIndex = -1 Then
                MsgBox "Seleccione la Condición del Documento...", vbCritical, sMensaje
                CboCondicion.SetFocus
                Exit Sub
            End If
        End If
        
        sTituloG = ""
        sTituloG = sTituloG & "Familia: Todas"
        sTituloG = sTituloG & " / Sub-Familia: Todas"
        
        If OptTipoFecha(0).value = True Then
           strTipoFiltro = "Filtrado por: Fecha de Ingreso"
        Else
            If OptTipoFecha(1).value Then
                strTipoFiltro = "Filtrado por: Fecha de Emisión"
            Else
                strTipoFiltro = "Filtrado por: Fecha de Proceso"
            End If
        End If
        sTitulo = "Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.value, "hh:mm") & " al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.value, "hh:mm")
        
        If Index <> 3 Then
            For w = 1 To lstServidores.ListItems.Count
                If lstServidores.ListItems.Item(w).Checked = True Then
                    lSw = True
                    Exit For
                End If
            Next
            If lSw = False Then
                MsgBox "Seleccione una Localidad!", vbCritical, sMensaje
                Exit Sub
            End If
        End If
        
        Screen.MousePointer = vbHourglass
        GeneraDetalle
        
        If rsReporte.EOF Then
            Screen.MousePointer = vbDefault
            MsgBox "No hay datos que mostrar en este rango de fechas...", vbExclamation, "Mensaje del Sistema"
            Exit Sub
        End If
        
        Select Case Index
               Case Is = 0 'Preliminar
                    frmEmite.CRViewer.DisplayGroupTree = False
                    If optFormato(0).value Then  'Por Tipo de Documento
                        CrtReporteTD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        CrtReporteTD.Database.SetDataSource rsReporte
                        CrtReporteTD.DiscardSavedData
                        CrtReporteTD.ReportTitle = sTitulo
                        CrtReporteTD.txtTitulo.SetText sTituloG
                        CrtReporteTD.txtCondicion.SetText "Condición : " & IIf(CboCondicion.ListIndex = -1, "Todos", Choose(CboCondicion.ListIndex + 1, "Pendientes", "Amortizados", "Cancelados"))
                        CrtReporteTD.txtFiltrado.SetText strTipoFiltro
                        CrtReporteTD.TxtRSocial.SetText sRazonComercial
                        CrtReporteTD.Text23.SetText "Expresado en " & Calcular("Select tMonedaN As Codigo From tParametro", Cn)
                        CrtReporteTD.Text9.SetText IIf(sImpuesto1 = "", "No Definido", sImpuesto1)
                        CrtReporteTD.Text10.SetText IIf(sImpuesto2 = "", "No Definido", sImpuesto2)
                        CrtReporteTD.Text11.SetText IIf(sImpuesto3 = "", "No Definido", sImpuesto3)
                        frmEmite.CRViewer.ReportSource = CrtReporteTD
                        
                    ElseIf optFormato(1).value Then  'Por Tipo de Producto
                        CrtReporteTP.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        CrtReporteTP.Database.SetDataSource rsReporte
                        CrtReporteTP.DiscardSavedData
                        CrtReporteTP.ReportTitle = sTitulo
                        CrtReporteTD.txtCondicion.SetText "Condición : " & IIf(CboCondicion.ListIndex = -1, "Todos", Choose(CboCondicion.ListIndex + 1, "Pendientes", "Amortizados", "Cancelados"))
                        CrtReporteTP.txtTitulo.SetText sTituloG
                        CrtReporteTP.txtFiltrado.SetText strTipoFiltro
                        CrtReporteTP.TxtRSocial.SetText sRazonComercial
                        CrtReporteTP.Text23.SetText "Expresado en " & Calcular("Select tMonedaN As Codigo From tParametro", Cn)
                        CrtReporteTP.Text9.SetText IIf(sImpuesto1 = "", "No Definido", sImpuesto1)
                        CrtReporteTP.Text10.SetText IIf(sImpuesto2 = "", "No Definido", sImpuesto2)
                        CrtReporteTP.Text11.SetText IIf(sImpuesto3 = "", "No Definido", sImpuesto3)
                        frmEmite.CRViewer.ReportSource = CrtReporteTP
                        
                    Else 'Registro de Compras
                        CrtReporteRC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        CrtReporteRC.Database.SetDataSource rsReporte
                        CrtReporteRC.DiscardSavedData
                        CrtReporteRC.ReportTitle = sTitulo
                        CrtReporteRC.txtTitulo.SetText sTituloG
                        CrtReporteTD.txtCondicion.SetText "Condición : " & IIf(CboCondicion.ListIndex = -1, "Todos", Choose(CboCondicion.ListIndex + 1, "Pendientes", "Amortizados", "Cancelados"))
                        CrtReporteRC.txtFiltrado.SetText strTipoFiltro
                        CrtReporteRC.TxtRSocial.SetText sRazonComercial
                        CrtReporteRC.Text23.SetText "Expresado en " & Calcular("Select tMonedaN As Codigo From tParametro", Cn)
                        CrtReporteRC.Text9.SetText IIf(sImpuesto1 = "", "No Definido", sImpuesto1)
                        CrtReporteRC.Text10.SetText IIf(sImpuesto2 = "", "No Definido", sImpuesto2)
                        CrtReporteRC.Text11.SetText IIf(sImpuesto3 = "", "No Definido", sImpuesto3)
                        CrtReporteRC.PaperOrientation = 2
                        frmEmite.CRViewer.ReportSource = CrtReporteRC
                    End If
                    frmEmite.CRViewer.ViewReport
                    frmEmite.Show vbModal
                    
               Case Is = 1 'Print
                    If optFormato(0).value Then  'Por Tipo de Documento
                        CrtReporteTD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        CrtReporteTD.Database.SetDataSource rsReporte
                        CrtReporteTD.ReportTitle = sTitulo
                        CrtReporteTD.txtFiltrado.SetText strTipoFiltro
                        CrtReporteTD.TxtRSocial.SetText sRazonComercial
                        CrtReporteTD.Text23.SetText "Expresado en " & Calcular("Select tMonedaN As Codigo From tParametro", Cn)
                        CrtReporteTD.Text9.SetText IIf(sImpuesto1 = "", "No Definido", sImpuesto1)
                        CrtReporteTD.Text10.SetText IIf(sImpuesto2 = "", "No Definido", sImpuesto2)
                        CrtReporteTD.Text11.SetText IIf(sImpuesto3 = "", "No Definido", sImpuesto3)
                        CrtReporteTD.PaperOrientation = 2
                        CrtReporteTD.DiscardSavedData
                        CrtReporteTD.PrintOut
                        
                    ElseIf optFormato(1).value Then  'Por Tipo de Producto
                        CrtReporteTP.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        CrtReporteTP.Database.SetDataSource rsReporte
                        CrtReporteTP.ReportTitle = sTitulo
                        CrtReporteTP.txtFiltrado.SetText strTipoFiltro
                        CrtReporteTP.TxtRSocial.SetText sRazonComercial
                        CrtReporteTP.Text23.SetText "Expresado en " & Calcular("Select tMonedaN As Codigo From tParametro", Cn)
                        CrtReporteTP.Text9.SetText IIf(sImpuesto1 = "", "No Definido", sImpuesto1)
                        CrtReporteTP.Text10.SetText IIf(sImpuesto2 = "", "No Definido", sImpuesto2)
                        CrtReporteTP.Text11.SetText IIf(sImpuesto3 = "", "No Definido", sImpuesto3)
                        CrtReporteTP.PaperOrientation = 2
                        CrtReporteTP.DiscardSavedData
                        CrtReporteTP.PrintOut
                    
                    Else 'Registro de compras
                        CrtReporteRC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        CrtReporteRC.Database.SetDataSource rsReporte
                        CrtReporteRC.ReportTitle = sTitulo
                        CrtReporteRC.txtFiltrado.SetText strTipoFiltro
                        CrtReporteRC.TxtRSocial.SetText sRazonComercial
                        CrtReporteRC.Text23.SetText "Expresado en " & Calcular("Select tMonedaN As Codigo From tParametro", Cn)
                        CrtReporteRC.Text9.SetText IIf(sImpuesto1 = "", "No Definido", sImpuesto1)
                        CrtReporteRC.Text10.SetText IIf(sImpuesto2 = "", "No Definido", sImpuesto2)
                        CrtReporteRC.Text11.SetText IIf(sImpuesto3 = "", "No Definido", sImpuesto3)
                        CrtReporteRC.PaperOrientation = 2
                        CrtReporteRC.DiscardSavedData
                        CrtReporteRC.PrintOut
                        
                    End If
                    
               Case Is = 2 'Excel
                    If optFormato(0).value Then  'Por Tipo de Documento
                        CrtReporteTD.ExportOptions.FormatType = crEFTExcel50
                        CrtReporteTD.ExportOptions.DestinationType = crEDTDiskFile
                        CrtReporteTD.Database.SetDataSource rsReporte
                        CrtReporteTD.ReportTitle = sTitulo
                        CrtReporteTD.txtFiltrado.SetText strTipoFiltro
                        CrtReporteTD.TxtRSocial.SetText sRazonComercial
                        CrtReporteTD.Text23.SetText "Expresado en " & Calcular("Select tMonedaN As Codigo From tParametro", Cn)
                        CrtReporteTD.Text9.SetText IIf(sImpuesto1 = "", "No Definido", sImpuesto1)
                        CrtReporteTD.Text10.SetText IIf(sImpuesto2 = "", "No Definido", sImpuesto2)
                        CrtReporteTD.Text11.SetText IIf(sImpuesto3 = "", "No Definido", sImpuesto3)
                        CrtReporteTD.PaperOrientation = 2
                        CrtReporteTD.DiscardSavedData
                        cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                        cmdSave.ShowSave
                        CrtReporteTD.ExportOptions.DiskFileName = cmdSave.FileName
                        If Trim(CrtReporteTD.ExportOptions.DiskFileName) <> "" Then CrtReporteTD.Export False
                        
                    ElseIf optFormato(1).value Then  'Por Tipo de Producto
                        CrtReporteTP.ExportOptions.FormatType = crEFTExcel50
                        CrtReporteTP.ExportOptions.DestinationType = crEDTDiskFile
                        CrtReporteTP.Database.SetDataSource rsReporte
                        CrtReporteTP.ReportTitle = sTitulo
                        CrtReporteTP.txtFiltrado.SetText strTipoFiltro
                        CrtReporteTP.TxtRSocial.SetText sRazonComercial
                        CrtReporteTP.Text23.SetText "Expresado en " & Calcular("Select tMonedaN As Codigo From tParametro", Cn)
                        CrtReporteTP.Text9.SetText IIf(sImpuesto1 = "", "No Definido", sImpuesto1)
                        CrtReporteTP.Text10.SetText IIf(sImpuesto2 = "", "No Definido", sImpuesto2)
                        CrtReporteTP.Text11.SetText IIf(sImpuesto3 = "", "No Definido", sImpuesto3)
                        CrtReporteTP.PaperOrientation = 2
                        CrtReporteTP.DiscardSavedData
                        cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                        cmdSave.ShowSave
                        CrtReporteTP.ExportOptions.DiskFileName = cmdSave.FileName
                        If Trim(CrtReporteTP.ExportOptions.DiskFileName) <> "" Then CrtReporteTP.Export False
                    
                    Else
                        CrtReporteRC.ExportOptions.FormatType = crEFTExcel50
                        CrtReporteRC.ExportOptions.DestinationType = crEDTDiskFile
                        CrtReporteRC.Database.SetDataSource rsReporte
                        CrtReporteRC.ReportTitle = sTitulo
                        CrtReporteRC.txtFiltrado.SetText strTipoFiltro
                        CrtReporteRC.TxtRSocial.SetText sRazonComercial
                        CrtReporteRC.Text23.SetText "Expresado en " & Calcular("Select tMonedaN As Codigo From tParametro", Cn)
                        CrtReporteRC.Text9.SetText IIf(sImpuesto1 = "", "No Definido", sImpuesto1)
                        CrtReporteRC.Text10.SetText IIf(sImpuesto2 = "", "No Definido", sImpuesto2)
                        CrtReporteRC.Text11.SetText IIf(sImpuesto3 = "", "No Definido", sImpuesto3)
                        CrtReporteRC.PaperOrientation = 2
                        CrtReporteRC.DiscardSavedData
                        cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                        cmdSave.ShowSave
                        CrtReporteRC.ExportOptions.DiskFileName = cmdSave.FileName
                        If Trim(CrtReporteRC.ExportOptions.DiskFileName) <> "" Then CrtReporteRC.Export False
                    
                    End If
                    
        End Select
        If sNombre <> "" Then Cn.Execute "Drop Table " & sNombre
        Screen.MousePointer = vbDefault
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Centrar Me
   
    LlenaCombo
   
    dtpFecIni.value = "01/" & str(Month(Date)) & "/" & str(Year(Date))
    dtpFecFin.value = Date
    
    obtieneListaServidores
End Sub

