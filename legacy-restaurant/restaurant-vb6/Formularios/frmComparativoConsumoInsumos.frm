VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmComparativoConsumoInsumos 
   Caption         =   "Comparativo de Consumo"
   ClientHeight    =   5925
   ClientLeft      =   1695
   ClientTop       =   345
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   10875
   Begin VB.Frame Frame4 
      Caption         =   "Formato"
      ForeColor       =   &H00800000&
      Height          =   675
      Left            =   120
      TabIndex        =   28
      Top             =   2985
      Width           =   6090
      Begin VB.OptionButton optformato 
         Caption         =   "Comparativo"
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   2
         Left            =   2715
         TabIndex        =   30
         Top             =   270
         Width           =   2475
      End
      Begin VB.OptionButton optformato 
         Caption         =   "Agrupado"
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   1
         Left            =   660
         TabIndex        =   29
         Top             =   270
         Value           =   -1  'True
         Width           =   1500
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
      Left            =   7590
      Picture         =   "frmComparativoConsumoInsumos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5250
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
      Left            =   6150
      Picture         =   "frmComparativoConsumoInsumos.frx":00F2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5250
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
      Left            =   4710
      Picture         =   "frmComparativoConsumoInsumos.frx":01BA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5250
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
      Left            =   3270
      Picture         =   "frmComparativoConsumoInsumos.frx":09DC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5250
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
      Height          =   5160
      Left            =   45
      TabIndex        =   4
      Top             =   0
      Width           =   10770
      Begin VB.Frame Frame2 
         Caption         =   " Tipo de Reporte "
         ForeColor       =   &H00800000&
         Height          =   675
         Left            =   75
         TabIndex        =   31
         Top             =   3690
         Width           =   6090
         Begin VB.OptionButton optOpcion 
            Caption         =   "Costo"
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   1
            Left            =   2715
            TabIndex        =   33
            Top             =   270
            Width           =   2475
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Cantidad"
            ForeColor       =   &H00000080&
            Height          =   240
            Index           =   0
            Left            =   645
            TabIndex        =   32
            Top             =   270
            Value           =   -1  'True
            Width           =   1500
         End
      End
      Begin VB.CheckBox Chkarea 
         Caption         =   "Todas"
         Height          =   195
         Left            =   5085
         TabIndex        =   25
         Top             =   2595
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox ChkSubFamilia 
         Caption         =   "Todas"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5085
         TabIndex        =   19
         Top             =   1710
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox ChkFamilia 
         Caption         =   "Todas"
         Height          =   195
         Left            =   5055
         TabIndex        =   18
         Top             =   1275
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.TextBox TxtCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2115
         Width           =   885
      End
      Begin VB.TextBox TxtArticulo 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2415
         TabIndex        =   16
         Top             =   2115
         Width           =   2580
      End
      Begin VB.CheckBox ChkArticulo 
         Caption         =   "Todos"
         Height          =   210
         Left            =   5085
         TabIndex        =   15
         Top             =   2160
         Value           =   1  'Checked
         Width           =   900
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
         TabIndex        =   5
         Top             =   0
         Width           =   4545
         Begin VB.CheckBox chkTodo 
            Caption         =   "Seleccionar todo"
            Height          =   300
            Left            =   2880
            TabIndex        =   6
            Top             =   4830
            Width           =   1575
         End
         Begin MSComctlLib.ListView lstServidores 
            Height          =   4605
            Left            =   90
            TabIndex        =   7
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
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   7170
         Top             =   4725
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "Xls"
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   330
         Left            =   1530
         TabIndex        =   8
         Top             =   780
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Format          =   55050241
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   330
         Left            =   1530
         TabIndex        =   9
         Top             =   330
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Format          =   55050241
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   330
         Left            =   3150
         TabIndex        =   10
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
         Format          =   55050243
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   330
         Left            =   3150
         TabIndex        =   11
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
         Format          =   55050243
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
      End
      Begin MSDataListLib.DataCombo CboFamilia 
         Height          =   315
         Left            =   1500
         TabIndex        =   20
         Top             =   1215
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   -2147483633
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
      Begin MSDataListLib.DataCombo CboSubFamilia 
         Height          =   315
         Left            =   1500
         TabIndex        =   21
         Top             =   1665
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   -2147483633
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
      Begin MSDataListLib.DataCombo dtcArea 
         Height          =   315
         Left            =   1515
         TabIndex        =   26
         Top             =   2565
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   -2147483633
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
         Caption         =   "Área:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   27
         Top             =   2580
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Familia:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   24
         Top             =   1695
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Familia:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   930
         TabIndex        =   23
         Top             =   1230
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Artículo:"
         Height          =   195
         Index           =   4
         Left            =   855
         TabIndex        =   22
         Top             =   2175
         Width           =   600
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
         TabIndex        =   13
         Top             =   840
         Width           =   915
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
         TabIndex        =   12
         Top             =   390
         Width           =   990
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
      Left            =   1830
      Picture         =   "frmComparativoConsumoInsumos.frx":0F0E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5250
      Width           =   1455
   End
End
Attribute VB_Name = "frmComparativoConsumoInsumos"
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
'Dim CrtReporteRC  As New CrtRegComprasIntegrado
Dim CrtReporteT1  As New dsrCompConsumoInsumos
Dim CrtReporteT2  As New dsrCompConsumoInsumosValor
Dim CrtReporteCompT1  As New dsrComparativoConsumoInsumos
Dim CrtReporteCompT2  As New dsrComparativoConsumoInsumosValor
'Dim CrtReporteTP  As New CrtComprasxTipoProductoIntegrado

Dim RsFamilia     As Recordset
Dim RsSubFamilia  As Recordset
Dim RsArea        As Recordset
Dim RsEstado      As Recordset
Dim RsTemporal    As Recordset

Dim tCondicion As String

Sub CreaTemporal()

        sNombre = dbTemporal(sUsuario, 13, "tLocalidad", "nVarChar(100)", _
                                           "tcodigofamilia", "nVarChar(2)", _
                                           "familia", "nVarChar(250)", _
                                           "tcodigosubfamilia", "nVarChar(4)", _
                                           "subFamilia", "nVarChar(250)", _
                                           "tcodigoproducto", "nVarChar(7)", _
                                           "tdetallado", "nVarChar(250)", _
                                           "tunidadsalida", "nVarChar(3)", _
                                           "unidadsalida", "nVarChar(100)", _
                                           "ConsumoReal", "Float", _
                                           "ConsumoRealValor", "Float", _
                                           "ConsumoPrevisto", "Float", _
                                           "ConsumoPrevistoValor", "Float")
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
        .ColumnHeaders.Add , , "IpINF", 0
        .ColumnHeaders.Add , , "BdINF", 0
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
            Item.SubItems(6) = rsS.Fields(3) 'Ip Almacen
            Item.SubItems(7) = rsS.Fields(4) 'Base de Datos Almacen
            rsS.MoveNext
        Next X
    End With
End Sub

Public Sub GeneraDetalle()
    Dim RsLista As ADODB.Recordset
    Dim oComando As clsComando
    Dim w, cl
    Dim conectado As Boolean
    Dim sector As String, llocal As String, Localidad As String, sOrden As String
    
    
    CreaTemporal
        
    
    Set oComando = New clsComando

    If sNombre <> "" Then Cn.Execute "Delete From " & sNombre

    For w = 1 To lstServidores.ListItems.Count
        If lstServidores.ListItems.Item(w).Checked = True Then
            conectado = validaConexion(lstServidores.ListItems.Item(w).SubItems(6), lstServidores.ListItems.Item(w).SubItems(7))
            If conectado = False Then
                MsgBox "Error Fatal: No se pudo conectar con  " & lstServidores.ListItems.Item(w).SubItems(2), vbCritical, sMensaje
                lstServidores.ListItems.Item(w).Checked = False
                GoTo sigue
            End If

                If Not oComando.CreateCmdSp("sp_ComparativoConsumo", devuelveConexion(lstServidores.ListItems.Item(w).SubItems(6), lstServidores.ListItems.Item(w).SubItems(7))) Then
                    Set oComando = Nothing
                    Exit Sub
                End If
            
            oComando.CreateParameter "@BDALMACEN", adVarChar, adParamInput, 25, lstServidores.ListItems.Item(w).SubItems(5)
            oComando.CreateParameter "@finicio", adVarChar, adParamInput, 20, Format(dtpFecIni.value, "yyyy/MM/dd") & " " & Format(dtpHorIni.value, "HH:mm:ss")
            oComando.CreateParameter "@ffinal", adVarChar, adParamInput, 20, Format(dtpFecFin.value, "yyyy/MM/dd") & " " & Format(dtpHorFin.value, "HH:mm:ss")
            oComando.CreateParameter "@vch_SubArea", adVarChar, adParamInput, 3, IIf(dtcArea.BoundText = "", "", dtcArea.BoundText)
            oComando.CreateParameter "@vch_Condicion", adVarChar, adParamInput, 3000, tCondicion
            If Not oComando.GetParamOK Then
                Set oComando = Nothing
                Exit Sub
            End If
            cl = lstServidores.ListItems.Item(w)
            sector = lstServidores.ListItems.Item(w).SubItems(1)
            llocal = lstServidores.ListItems.Item(w).SubItems(2)
            Set RsDatos = oComando.GetSP()
            If RsDatos.RecordCount > 0 Then
            RsDatos.MoveFirst
            With RsDatos
                While Not RsDatos.EOF

                    
                        Isql = "INSERT INTO " & sNombre & "(tLocalidad, tcodigofamilia, familia, tcodigosubfamilia,subfamilia, tcodigoproducto, tdetallado , tunidadsalida ,unidadsalida, ConsumoReal, ConsumoRealValor, ConsumoPrevisto , ConsumoPrevistoValor) " & _
                               "VALUES('" & llocal & "','" & !tcodigofamilia & "','" & !familia & "','" & !tcodigosubfamilia & "', '" & !subFamilia & "','" & !tCodigoProducto & "','" & !tDetallado & "','" & !tunidadsalida & "','" & !unidadsalida & "'," & !ConsumoRealCant & "," & !ConsumoRealValor & "," & !ConsumoPrevisto & "," & !ConsumoPrevistoValor & ")"
                    
                    Cn.Execute Isql
                    RsDatos.MoveNext
                Wend
            End With
            End If
        End If
sigue:
    Next w
    
        Isql = " Select tLocalidad, tcodigofamilia, familia, tcodigosubfamilia, subfamilia, " & _
               " tcodigoproducto, tdetallado , tunidadsalida ,unidadsalida, ConsumoReal, " & _
               " ConsumoRealValor, ConsumoPrevisto , ConsumoPrevistoValor from " & sNombre & " order by tLocalidad, tcodigoproducto"

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
        .ColumnHeaders.Add , , "IpINF", 0
        .ColumnHeaders.Add , , "BdINF", 0
        .View = lvwReport
        .GridLines = True
    
        rsS.MoveFirst
        For X = 0 To rsS.RecordCount - 1
            Set Item = .ListItems.Add(, , rsS.Fields(0))
            Item.SubItems(1) = rsS.Fields(1) 'Sector
            Item.SubItems(2) = rsS.Fields(2) 'Local
            Item.SubItems(3) = rsS.Fields(5) 'Estado
            Item.SubItems(4) = rsS.Fields(6) 'Ip Almacen
            Item.SubItems(5) = rsS.Fields(7) 'Base de Datos Almacen
            Item.SubItems(6) = rsS.Fields(3) 'Ip Almacen
            Item.SubItems(7) = rsS.Fields(4) 'Base de Datos Almacen
        rsS.MoveNext
        Next X
    End With
End Sub


Sub LlenaCombo()
    If Not CnAlmacen Is Nothing Then
        Isql = "Select Codigo, Descripcion from vFamilia where lActivo=1 Order By 2"
        Llena_Combo Isql, CboFamilia, RsFamilia, "Codigo", "Descripcion"
        
        Isql = "Select Codigo, Descripcion from vsubFamilia where lActivo=1 Order By 2"
        Llena_Combo Isql, CboSubFamilia, RsSubFamilia, "Codigo", "Descripcion"
        
        Isql = "Select Codigo, Descripcion from vArea where lActivo=1 Order By 2"
        Llena_Combo Isql, dtcArea, RsArea, "Codigo", "Descripcion"
    
    End If
End Sub

Private Sub CboFamilia_Change()
    CboSubFamilia.BoundText = ""
End Sub

Private Sub CboFamilia_LostFocus()
    Isql = "Select Codigo, Descripcion From vSubFamilia Where lActivo=1 And tFamilia='" & CboFamilia.BoundText & "'"
    Llena_Combo Isql, CboSubFamilia, RsSubFamilia, "Codigo", "Descripcion"
End Sub

Private Sub Chkarea_Click()
    If Chkarea.value = 1 Then
        dtcArea = ""
        dtcArea.BackColor = &H8000000F
        dtcArea.Enabled = False
    Else
        dtcArea.Enabled = True
        dtcArea.BackColor = &HF0EADB
        dtcArea.SetFocus
    End If
End Sub

Private Sub ChkArticulo_Click()
    If ChkArticulo.value = 1 Then
        TxtCodigo.Text = ""
        TxtArticulo = ""
        TxtArticulo.BackColor = &H8000000F
        TxtArticulo.Enabled = False
    Else
        TxtArticulo.Enabled = True
        TxtArticulo.BackColor = &HF0EADB
        TxtArticulo.SetFocus
    End If
End Sub

Private Sub ChkFamilia_Click()
    If ChkFamilia.value = 1 Then
       CboFamilia.Enabled = False
       CboFamilia = ""
       CboFamilia.BackColor = &H8000000F
       ChkSubFamilia.value = 1
       ChkSubFamilia.Enabled = False
       CboSubFamilia.Enabled = False
       CboSubFamilia = ""
       CboSubFamilia.BackColor = &H8000000F
       
       ChkArticulo.Enabled = True
       TxtArticulo.Enabled = True
       TxtArticulo.BackColor = &H8000000F
       
    Else
       ChkSubFamilia.Enabled = True
       CboFamilia.Enabled = True
       CboFamilia.BackColor = &HF0EADB
       
       ChkArticulo.Enabled = False
       ChkArticulo.value = 1
       TxtArticulo.Text = ""
       TxtArticulo.Enabled = False
       TxtArticulo.BackColor = &H8000000F
    End If
End Sub

Private Sub ChkSubFamilia_Click()
    If ChkSubFamilia.value = 1 Then
       CboSubFamilia.Enabled = False
       CboSubFamilia = ""
       CboSubFamilia.BackColor = &H8000000F
    Else
       CboSubFamilia.Enabled = True
       CboSubFamilia.BackColor = &HF0EADB
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
        
        tCondicion = ""
        If ChkFamilia.value = 0 Then
            If CboFamilia.BoundText = "" Then
                MsgBox "Seleccione la Familia...", vbCritical, sMensaje
                CboFamilia.SetFocus
                Exit Sub
            End If
            tCondicion = tCondicion & " tcodigofamilia='" & CboFamilia.BoundText & "'"
        End If
        If ChkSubFamilia.value = 0 Then
            If CboSubFamilia.BoundText = "" Then
                MsgBox "Seleccione la Sub-Familia...", vbCritical, sMensaje
                CboSubFamilia.SetFocus
                Exit Sub
            End If
            tCondicion = tCondicion & " and tcodigosubfamilia='" & CboSubFamilia.BoundText & "'"
        End If

        If ChkArticulo.value = 0 Then
            If TxtCodigo.Text = "" Then
                MsgBox "Seleccione un Articulo...", vbCritical, sMensaje
                TxtArticulo.SetFocus
                Exit Sub
            End If
            tCondicion = tCondicion & " and tcodigoproducto=" & Me.TxtCodigo.Text
        End If
        If Left(tCondicion, 4) = " and" Then
            tCondicion = Right(tCondicion, Len(tCondicion) - 4)
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
                    If optformato(1).value Then
                        If optOpcion(0).value Then  'Por Tipo de Documento
                            CrtReporteT1.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                            CrtReporteT1.Database.SetDataSource rsReporte
                            CrtReporteT1.DiscardSavedData
                            CrtReporteT1.ReportTitle = sTitulo
                            frmEmite.CRViewer.ReportSource = CrtReporteT1
                        Else
                            CrtReporteT2.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                            CrtReporteT2.Database.SetDataSource rsReporte
                            CrtReporteT2.DiscardSavedData
                            CrtReporteT2.ReportTitle = sTitulo
                            frmEmite.CRViewer.ReportSource = CrtReporteT2
                        End If
                    Else
                        If optOpcion(0).value Then  'Por Tipo de Documento
                            CrtReporteCompT1.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                            CrtReporteCompT1.Database.SetDataSource rsReporte
                            CrtReporteCompT1.DiscardSavedData
                            CrtReporteCompT1.ReportTitle = sTitulo & " - " & "Por Cantidad"
                            CrtReporteCompT1.PaperOrientation = 2
                            frmEmite.CRViewer.ReportSource = CrtReporteCompT1
                        Else
                            CrtReporteCompT2.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                            CrtReporteCompT2.Database.SetDataSource rsReporte
                            CrtReporteCompT2.DiscardSavedData
                            CrtReporteCompT2.ReportTitle = sTitulo & " - " & "Por Valor"
                            CrtReporteCompT1.PaperOrientation = 2
                            frmEmite.CRViewer.ReportSource = CrtReporteCompT2
                        End If
                    End If
                    frmEmite.CRViewer.ViewReport
                    frmEmite.Show vbModal
                    
               Case Is = 1 'Print
                    If optformato(1).value Then
                        If optOpcion(0).value Then  'Por Tipo de Documento
                            CrtReporteT1.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                            CrtReporteT1.Database.SetDataSource rsReporte
                            CrtReporteT1.ReportTitle = sTitulo
                            CrtReporteT1.PaperOrientation = 1
                            CrtReporteT1.DiscardSavedData
                            CrtReporteT1.PrintOut
                        Else
                            CrtReporteT2.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                            CrtReporteT2.Database.SetDataSource rsReporte
                            CrtReporteT2.ReportTitle = sTitulo
                            CrtReporteT2.PaperOrientation = 1
                            CrtReporteT2.DiscardSavedData
                            CrtReporteT2.PrintOut
                        End If
                    Else
                        If optOpcion(0).value Then  'Por Tipo de Documento
                            CrtReporteCompT1.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                            CrtReporteCompT1.Database.SetDataSource rsReporte
                            CrtReporteCompT1.ReportTitle = sTitulo & " - " & "Por Cantidad"
                            CrtReporteCompT1.PaperOrientation = 2
                            CrtReporteCompT1.DiscardSavedData
                            CrtReporteCompT1.PrintOut
                        Else
                            CrtReporteCompT2.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                            CrtReporteCompT2.Database.SetDataSource rsReporte
                            CrtReporteCompT2.ReportTitle = sTitulo & " - " & "Por Valor"
                            CrtReporteCompT2.PaperOrientation = 2
                            CrtReporteCompT2.DiscardSavedData
                            CrtReporteCompT2.PrintOut
                        End If
                    End If
               Case Is = 2 'Excel
                    If optformato(1).value Then
                        If optOpcion(0).value Then  'Por Tipo de Documento
                            CrtReporteT1.ExportOptions.FormatType = crEFTExcel50
                            CrtReporteT1.ExportOptions.DestinationType = crEDTDiskFile
                            CrtReporteT1.Database.SetDataSource rsReporte
                            CrtReporteT1.ReportTitle = sTitulo
                            CrtReporteT1.PaperOrientation = 1
                            CrtReporteT1.DiscardSavedData
                            cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                            cmdSave.ShowSave
                            CrtReporteT1.ExportOptions.DiskFileName = cmdSave.FileName
                            If Trim(CrtReporteT1.ExportOptions.DiskFileName) <> "" Then CrtReporteT1.Export False
                        Else
                            CrtReporteT2.ExportOptions.FormatType = crEFTExcel50
                            CrtReporteT2.ExportOptions.DestinationType = crEDTDiskFile
                            CrtReporteT2.Database.SetDataSource rsReporte
                            CrtReporteT2.ReportTitle = sTitulo
                            CrtReporteT2.PaperOrientation = 1
                            CrtReporteT2.DiscardSavedData
                            cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                            cmdSave.ShowSave
                            CrtReporteT2.ExportOptions.DiskFileName = cmdSave.FileName
                            If Trim(CrtReporteT2.ExportOptions.DiskFileName) <> "" Then CrtReporteT2.Export False
                        End If
                    Else
                        If optOpcion(0).value Then  'Por Tipo de Documento
                            CrtReporteCompT1.ExportOptions.FormatType = crEFTExcel50
                            CrtReporteCompT1.ExportOptions.DestinationType = crEDTDiskFile
                            CrtReporteCompT1.Database.SetDataSource rsReporte
                            CrtReporteCompT1.ReportTitle = sTitulo & " - " & "Por Cantidad"
                            CrtReporteCompT1.PaperOrientation = 2
                            CrtReporteCompT1.DiscardSavedData
                            cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                            cmdSave.ShowSave
                            CrtReporteT1.ExportOptions.DiskFileName = cmdSave.FileName
                            If Trim(CrtReporteCompT1.ExportOptions.DiskFileName) <> "" Then CrtReporteCompT1.Export False
                        Else
                            CrtReporteCompT2.ExportOptions.FormatType = crEFTExcel50
                            CrtReporteCompT2.ExportOptions.DestinationType = crEDTDiskFile
                            CrtReporteCompT2.Database.SetDataSource rsReporte
                            CrtReporteCompT2.ReportTitle = sTitulo & " - " & "Por Valor"
                            CrtReporteCompT2.PaperOrientation = 2
                            CrtReporteCompT2.DiscardSavedData
                            cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                            cmdSave.ShowSave
                            CrtReporteCompT2.ExportOptions.DiskFileName = cmdSave.FileName
                            If Trim(CrtReporteCompT2.ExportOptions.DiskFileName) <> "" Then CrtReporteCompT2.Export False
                        End If
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

Private Sub TxtArticulo_Change()
    If Len(Trim(UCase(TxtArticulo))) = 1 Then
        Isql = "Select tProducto.tCodigoProducto As Codigo, tProducto.tDetallado As Descripcion, vUnidadMedida.Descripcion AS UnidadEntrada" & _
               " From tProducto Left Outer Join vUnidadMedida ON tProducto.tUnidadEntrada = vUnidadMedida.Codigo "
        Call ConfGrilla(3, frmBuscaAlmacen.grdGrilla, "Codigo", 2, "Codigo", 2300, 2, 0, "", _
                                                      "Insumo", 2, "Descripcion", 5000, 0, 0, "", _
                                                      "Unidad", 2, "UnidadEntrada", 1000, 0, 0, "")
               
        frmBuscaAlmacen.nPredeterm = 1
        frmBuscaAlmacen.txtCriterio = Trim(UCase(TxtArticulo))
        frmBuscaAlmacen.txtCriterio.SelStart = Len(Trim(frmBuscaAlmacen.txtCriterio))
        sCodigo = ""
        sDescrip = ""
        
        frmBuscaAlmacen.Show vbModal
       
        TxtCodigo.Text = sCodigo
        TxtArticulo.Text = sDescrip
    End If
End Sub
