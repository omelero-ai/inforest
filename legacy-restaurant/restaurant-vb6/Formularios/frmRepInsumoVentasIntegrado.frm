VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRepInsumoVentasIntegrado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paloteo de Insumos"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   ForeColor       =   &H8000000C&
   Icon            =   "frmRepInsumoVentasIntegrado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
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
      Left            =   5273
      Picture         =   "frmRepInsumoVentasIntegrado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
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
      Left            =   8183
      Picture         =   "frmRepInsumoVentasIntegrado.frx":10EC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5760
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
      Left            =   3818
      Picture         =   "frmRepInsumoVentasIntegrado.frx":11DE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
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
      Left            =   2363
      Picture         =   "frmRepInsumoVentasIntegrado.frx":1710
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
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
      Left            =   6728
      Picture         =   "frmRepInsumoVentasIntegrado.frx":1C42
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.Frame Frame5 
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
      Height          =   5715
      Left            =   7320
      TabIndex        =   30
      Top             =   0
      Width           =   4575
      Begin VB.CheckBox chkTodo 
         Caption         =   "Seleccionar Todo"
         Height          =   360
         Left            =   2880
         TabIndex        =   24
         Top             =   5280
         Width           =   1575
      End
      Begin MSComctlLib.ListView lstServidores 
         Height          =   5055
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   8916
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
      Height          =   5715
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   7275
      Begin VB.Frame Frame6 
         Caption         =   "Ranking"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   120
         TabIndex        =   33
         Top             =   3480
         Width           =   2250
         Begin VB.OptionButton optRanking 
            Caption         =   "Por Cantidad"
            Height          =   240
            Index           =   0
            Left            =   75
            TabIndex        =   18
            Top             =   330
            Width           =   1305
         End
         Begin VB.OptionButton optRanking 
            Caption         =   "Por Montos"
            Height          =   240
            Index           =   1
            Left            =   75
            TabIndex        =   19
            Top             =   675
            Value           =   -1  'True
            Width           =   1185
         End
         Begin VB.OptionButton optRanking 
            Caption         =   "Ambos"
            Height          =   240
            Index           =   2
            Left            =   75
            TabIndex        =   20
            Top             =   1020
            Width           =   1185
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Visualización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   2400
         TabIndex        =   32
         Top             =   1680
         Width           =   2250
         Begin VB.OptionButton optVisualizacion 
            Caption         =   "Por Familia"
            Height          =   240
            Index           =   1
            Left            =   75
            TabIndex        =   12
            Top             =   675
            Width           =   1665
         End
         Begin VB.OptionButton optVisualizacion 
            Caption         =   "Detallado"
            Height          =   240
            Index           =   0
            Left            =   75
            TabIndex        =   11
            Top             =   330
            Value           =   -1  'True
            Width           =   1305
         End
         Begin VB.OptionButton optVisualizacion 
            Caption         =   "Por Sub Familia"
            Height          =   240
            Index           =   2
            Left            =   75
            TabIndex        =   13
            Top             =   1020
            Width           =   1425
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Agrupado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   2400
         TabIndex        =   31
         Top             =   3480
         Width           =   2250
         Begin VB.OptionButton optTipo 
            Caption         =   "Por Sector"
            Height          =   240
            Index           =   0
            Left            =   75
            TabIndex        =   21
            Top             =   405
            Width           =   1155
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "Por Local"
            Height          =   240
            Index           =   1
            Left            =   75
            TabIndex        =   22
            Top             =   975
            Value           =   -1  'True
            Width           =   1155
         End
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
         Height          =   1665
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   2250
         Begin VB.OptionButton optOpcion 
            Caption         =   "Todos los Insumos"
            Height          =   300
            Index           =   1
            Left            =   75
            TabIndex        =   9
            Top             =   405
            Value           =   -1  'True
            Width           =   1995
         End
         Begin VB.OptionButton optOpcion 
            Caption         =   "Insumos de control diario"
            Height          =   300
            Index           =   2
            Left            =   75
            TabIndex        =   10
            Top             =   975
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Origen de Datos "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   4680
         TabIndex        =   26
         Top             =   1680
         Width           =   2490
         Begin VB.CheckBox chkPlato 
            Caption         =   "Platos de Venta"
            Height          =   195
            Left            =   75
            TabIndex        =   14
            Top             =   270
            Value           =   1  'Checked
            Width           =   1770
         End
         Begin VB.CheckBox chkCombo 
            Caption         =   "Combos"
            Height          =   195
            Left            =   75
            TabIndex        =   15
            Top             =   585
            Value           =   1  'Checked
            Width           =   1725
         End
         Begin VB.CheckBox chkPropiedad 
            Caption         =   "Propiedades de los Platos"
            Height          =   195
            Left            =   75
            TabIndex        =   16
            Top             =   900
            Value           =   1  'Checked
            Width           =   2130
         End
         Begin VB.CheckBox chkPCombo 
            Caption         =   "Propiedades de los Combos"
            Height          =   195
            Left            =   75
            TabIndex        =   17
            Top             =   1200
            Value           =   1  'Checked
            Width           =   2265
         End
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   345
         Left            =   1830
         TabIndex        =   7
         Top             =   975
         Width           =   1785
         _ExtentX        =   3149
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
         Format          =   51118081
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   345
         Left            =   1830
         TabIndex        =   5
         Top             =   330
         Width           =   1785
         _ExtentX        =   3149
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
         Format          =   51118081
         CurrentDate     =   37539.2083333333
      End
      Begin MSComCtl2.DTPicker dtpHorIni 
         Height          =   375
         Left            =   3825
         TabIndex        =   6
         Top             =   330
         Width           =   1710
         _ExtentX        =   3016
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
         Format          =   51118083
         UpDown          =   -1  'True
         CurrentDate     =   37539
      End
      Begin MSComCtl2.DTPicker dtpHorFin 
         Height          =   375
         Left            =   3825
         TabIndex        =   8
         Top             =   975
         Width           =   1710
         _ExtentX        =   3016
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
         Format          =   51118083
         UpDown          =   -1  'True
         CurrentDate     =   37541.9993055556
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
         TabIndex        =   29
         Top             =   330
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
         TabIndex        =   28
         Top             =   975
         Width           =   1545
      End
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   1170
      Top             =   4575
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepInsumoVentasIntegrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNombre     As String
Dim rsReporte   As Recordset
Dim RsArea      As Recordset
Dim RsProducto  As Recordset
Dim Resumido   As New dsrInsumosIntegrados
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

Private Sub cmdOpcion_Click(Index As Integer)
   Cn.Execute "delete from " & sNombre
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   familia = ""
   subFamilia = ""

   Area = ""
   sCriterio = ""
   sTitulo = ""
   sFiltro = ""
   
  Select Case Index
          Case Is = 0 ' Preview
               
                    Genera

               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If

            
                frmEmite.CRViewer.DisplayGroupTree = True
                'Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                Resumido.PaperOrientation = crLandscape
                
                frmEmite.CRViewer.ViewReport
                frmEmite.Show vbModal
          
          Case Is = 1 ' Imprimir

                    Genera
           
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
             
                        Resumido.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                        Resumido.PaperOrientation = crLandscape
                        Resumido.PrintOut

          
          Case Is = 2 ' Salir
               Unload Me
          
          Case Is = 3 ' Exportar
              
                    Genera

               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
                        Resumido.ExportOptions.FormatType = 21
                        Resumido.ExportOptions.DestinationType = 1
                        cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                        cmdSave.ShowSave
                        If cmdSave.FileName = "" Then
                           Exit Sub
                        End If
                        Resumido.ExportOptions.DiskFileName = cmdSave.FileName
                        Resumido.Export False
   End Select
End Sub

Private Sub dtpFecfin_LostFocus()
   If dtpFecIni.value > dtpFecFin.value Then
      MsgBox "Error en Rango de Fechas", vbCritical, sMensaje
      dtpFecFin.SetFocus
   End If
End Sub

Private Sub Form_Load()
    Centrar Me
      sNombre = dbTemporal(sCaja, 8, "sector", "nvarchar(50)", "codlocal", "nvarchar(2)", "local", "nvarchar(50)", "Familia", "nvarchar(350)", "Subfamilia", "nvarchar(350)", "Insumo", "nvarchar(350)", "nConsumo", "Float", "Total", "Float")
    dtpFecIni.value = Date
    dtpFecFin.value = Date
    obtieneListaServidores
End Sub

Private Sub Form_Unload(Cancel As Integer)
     Cn.Execute "drop table " & sNombre
   Set rsReporte = Nothing
End Sub

Public Sub Genera()
   On Error GoTo sigue
   Dim oComando As New clsComando
   Dim conectado As Boolean
   Dim conectadoAl As Boolean
   Dim RsDatos As New ADODB.Recordset
   Dim w, z As Long
   Dim bdAlmacen As String
   
   Dim cl, sector, llocal As String
   
   Set oComando = New clsComando
   Screen.MousePointer = vbHourglass

    fInicio = Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHorIni.value, "HH:mm")
    fFinal = Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHorFin.value, "HH:mm")
                    
     For w = 1 To lstServidores.ListItems.Count
        If lstServidores.ListItems.Item(w).Checked = True Then
                conectado = validaConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))
                'bdAlmacen = Calcular("select isnull(tbasedatosalm,'') as codigo from tlocal where tcodigolocal='" & lstServidores.ListItems.Item(w) & "' and tbasedatosinf='" & lstServidores.ListItems.Item(w).SubItems(5) & "'", Cn)
                bdAlmacen = Calcular("select isnull(BdAlm,'') as codigo from vLocalidades where Codigo='" & lstServidores.ListItems.Item(w) & "'", CnAlmacen)
                conectadoAl = validaConexion(lstServidores.ListItems.Item(w).SubItems(4), bdAlmacen)
                If conectado = False Or conectadoAl = False Then
                    MsgBox "Error Fatal: No se pudo conectar con  " & lstServidores.ListItems.Item(w).SubItems(2), vbCritical, sMensaje
                    lstServidores.ListItems.Item(w).Checked = False
                    GoTo sigue
                End If
                
                Set oComando = New clsComando
                If Not oComando.CreateCmdSp("spRep_PaloteoInsumoIntegrado", devuelveConexion(lstServidores.ListItems.Item(w).SubItems(4), lstServidores.ListItems.Item(w).SubItems(5))) Then
                        Set oComando = Nothing
                        Exit Sub
                End If
                            oComando.CreateParameter "@dbAlmacen", adVarChar, adParamInput, 50, bdAlmacen
                            oComando.CreateParameter "@flagPlato", adBoolean, adParamInput, 1, Me.chkPlato.value
                            oComando.CreateParameter "@flagCombo", adBoolean, adParamInput, 1, Me.chkCombo.value
                            oComando.CreateParameter "@flagPropiedad", adBoolean, adParamInput, 1, Me.chkPropiedad.value
                            oComando.CreateParameter "@flagPCombo", adBoolean, adParamInput, 1, Me.chkPCombo.value
                            oComando.CreateParameter "@flagTipo", adBoolean, adParamInput, 1, optOpcion(2).value
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
                          RsDatos.MoveFirst
                          For z = 0 To RsDatos.RecordCount - 1
                                          Isql = "insert into " & sNombre & " values ( '" & sector & "','" & cl & "','" & llocal & "','" & RsDatos.Fields(0) & "','" & RsDatos.Fields(1) & "','" & RsDatos.Fields(2) & "'," & RsDatos.Fields(3) & "," & RsDatos.Fields(4) & ")"
                                         'Debug.Print Isql
                                          Cn.Execute Isql
                            RsDatos.MoveNext
                        Next z
                        
                 End If
        End If
sigue:
    Next w
        
    If optTipo(0).value = True Then ' sector
        If optRanking(0).value = True Then 'x cantidad
            If optVisualizacion(0).value = True Then
                Isql = "select sector as local,insumo, sum(nConsumo) as total from " & sNombre & " group by sector, insumo order by 1,2"
            ElseIf optVisualizacion(1).value = True Then
                Isql = "select sector as local,familia, sum(nConsumo) as total from " & sNombre & " group by sector, familia order by 1,2"
            ElseIf optVisualizacion(2).value = True Then
                Isql = "select sector as local,subfamilia, sum(nConsumo) as total from " & sNombre & " group by sector, subfamilia order by 1,2"
            End If
            
        
        ElseIf optRanking(1).value = True Then ' x montos
            If optVisualizacion(0).value = True Then
                Isql = "select sector as local,insumo, sum(total) as total from " & sNombre & " group by sector, insumo order by 1,2"
            ElseIf optVisualizacion(1).value = True Then
                Isql = "select sector as local,familia, sum(total) as total from " & sNombre & " group by sector, familia order by 1,2"
            ElseIf optVisualizacion(2).value = True Then
                Isql = "select sector as local,subfamilia, sum(total) as total from " & sNombre & " group by sector, subfamilia order by 1,2"
            End If
        ElseIf optRanking(2).value = True Then
            If optVisualizacion(0).value = True Then
                Isql = "select sector+'Mont' as local,insumo, sum(total) as total from " & sNombre & " group by sector, insumo Union select sector+'Cant' as local,insumo, sum(nConsumo) as total from " & sNombre & " group by sector, insumo  order by 1,2"
            ElseIf optVisualizacion(1).value = True Then
                Isql = "select sector+'Mont' as local,familia, sum(total) as total from " & sNombre & " group by sector, familia UNION select sector+'Cant' as local,familia, sum(nConsumo) as total from " & sNombre & " group by sector, familia  order by 1,2"
            ElseIf optVisualizacion(2).value = True Then
                Isql = "select sector+'Mont' as local,subfamilia, sum(total) as total from " & sNombre & " group by sector, subfamilia UNION select sector +'Cant' as local,subfamilia, sum(nConsumo) as total from " & sNombre & " group by sector, subfamilia order by 1,2"
            End If
        End If
    Else ' local
        If optRanking(0).value = True Then 'x cantidad
            If optVisualizacion(0).value = True Then
                Isql = "select local,insumo, sum(nConsumo) as total from " & sNombre & " group by local, insumo order by 1,2"
            ElseIf optVisualizacion(1).value = True Then
                Isql = "select local,familia, sum(nConsumo) as total from " & sNombre & " group by local, familia order by 1,2"
            ElseIf optVisualizacion(2).value = True Then
                Isql = "select local,subfamilia, sum(nConsumo) as total from " & sNombre & " group by local, subfamilia order by 1,2"
            End If
        ElseIf optRanking(1).value = True Then ' x montos
            If optVisualizacion(0).value = True Then
                Isql = "select local,insumo, sum(total) as total from " & sNombre & " group by local, insumo order by 1,2"
            ElseIf optVisualizacion(1).value = True Then
                Isql = "select local,familia, sum(total) as total from " & sNombre & " group by local, familia order by 1,2"
            ElseIf optVisualizacion(2).value = True Then
                Isql = "select local,subfamilia, sum(total) as total from " & sNombre & " group by local, subfamilia order by 1,2"
            End If
        ElseIf optRanking(2).value = True Then
            If optVisualizacion(0).value = True Then
                Isql = "select local+'Mont' as local,insumo, sum(total) as total from " & sNombre & " group by local, insumo Union select local+'Cant' as local,insumo, sum(nConsumo) as total from " & sNombre & " group by local, insumo  order by 1,2"
            ElseIf optVisualizacion(1).value = True Then
                Isql = "select local+'Mont' as local,familia, sum(total) as total from " & sNombre & " group by local, familia UNION select local+'Cant' as local,familia, sum(nConsumo) as total from " & sNombre & " group by local, familia  order by 1,2"
            ElseIf optVisualizacion(2).value = True Then
                Isql = "select local+'Mont' as local,subfamilia, sum(total) as total from " & sNombre & " group by local, subfamilia UNION select local +'Cant' as local,subfamilia, sum(nConsumo) as total from " & sNombre & " group by local, subfamilia order by 1,2"
            End If
        End If
    
    End If
    
    Set rsReporte = Lib.OpenRecordset(Isql, Cn)
    
    
        Resumido.DiscardSavedData
        Resumido.Database.SetDataSource rsReporte
        Resumido.Text1.SetText IIf(Me.optOpcion(1).value = True, optOpcion(1).Caption, optOpcion(2).Caption) & " . Visualización: " & IIf(Me.optRanking(0).value = True, optRanking(0).Caption, optRanking(1).Caption) & ". Agrupado: " & IIf(Me.optTipo(0).value = True, optTipo(0).Caption, optTipo(1).Caption)
        Resumido.Text2.SetText "Origen de Datos: " & IIf(chkPlato.value = 1, chkPlato.Caption, "") & IIf(chkCombo.value = 1, " -" & chkCombo.Caption, "") & IIf(chkPropiedad.value = 1, " -" & chkPropiedad.Caption, "") & IIf(chkPCombo.value = 1, " -" & chkPCombo.Caption, "")
        Resumido.Text20.SetText sRazonSocial
        '      Resumido.Text25.SetText localConectado
        Resumido.ReportTitle = " Del " & Format(dtpFecIni.value, "dd/mmm/yyyy") & " " & Format(dtpHorIni.value, "HH:mm") & " Al " & Format(dtpFecFin.value, "dd/mmm/yyyy") & " " & Format(dtpHorFin.value, "HH:mm")
        frmEmite.CRViewer.ReportSource = Resumido


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
       ' validaCantidadSeleccion
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
   Dim p, w As Integer
    Dim s As Integer
    If Item.SubItems(3) = "Desconectado" Then
        MsgBox "El enlace con el Servidor Seleccionado no es Posible", vbCritical, sMensaje
        Item.Checked = False
    End If
End Sub
Private Sub chkTodo_Click()
        activaCheck IIf(chkTodo.value = 1, True, False)
End Sub
