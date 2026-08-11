VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmFiltroProducto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Emite de Productos"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   Icon            =   "frmFiltroProducto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkPatio 
      Caption         =   "Imprimir Productos Patio de comidas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2880
      TabIndex        =   29
      Top             =   3120
      Width           =   3615
   End
   Begin VB.CheckBox chkVineta 
      Caption         =   "Imprimir Viñetas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   150
      TabIndex        =   28
      Top             =   3075
      Width           =   1785
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Viñeta"
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
      Height          =   615
      Index           =   4
      Left            =   4875
      Picture         =   "frmFiltroProducto.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Genera Viñeta"
      Top             =   3450
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
      Left            =   6350
      Picture         =   "frmFiltroProducto.frx":044E
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3450
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
      Left            =   1946
      Picture         =   "frmFiltroProducto.frx":0540
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3450
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
      Left            =   491
      Picture         =   "frmFiltroProducto.frx":0A72
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3450
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
      Left            =   3375
      Picture         =   "frmFiltroProducto.frx":0FA4
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3450
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame frameCanal 
      Caption         =   " Canal de Venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   2850
      TabIndex        =   20
      Top             =   1800
      Width           =   2745
      Begin MSDataListLib.DataCombo cboCanalVenta 
         Height          =   315
         Left            =   150
         TabIndex        =   21
         Top             =   570
         Width           =   2490
         _ExtentX        =   4392
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Filtrado por "
         Height          =   195
         Left            =   150
         TabIndex        =   22
         Top             =   300
         Width           =   825
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   " Estado Productos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   5640
      TabIndex        =   16
      Top             =   1800
      Width           =   2625
      Begin VB.OptionButton optEstado 
         Caption         =   "No Activos"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   19
         Top             =   855
         Width           =   2310
      End
      Begin VB.OptionButton optEstado 
         Caption         =   "Activos"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   18
         Top             =   585
         Width           =   2310
      End
      Begin VB.OptionButton optEstado 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Top             =   315
         Value           =   -1  'True
         Width           =   2310
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipo de Reporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   60
      TabIndex        =   12
      Top             =   1800
      Width           =   2745
      Begin VB.OptionButton optTipo 
         Caption         =   "Productos y Precios"
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   15
         Top             =   855
         Width           =   1785
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Visualizacion de Combinaciones"
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   14
         Top             =   585
         Width           =   2610
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Productos"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   13
         Top             =   315
         Value           =   -1  'True
         Width           =   1965
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " SubGrupo "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   5640
      TabIndex        =   10
      Top             =   30
      Width           =   2625
      Begin VB.CheckBox chkSubGrupo 
         Caption         =   "Todos"
         Height          =   285
         Left            =   180
         TabIndex        =   4
         Top             =   330
         Value           =   1  'Checked
         Width           =   855
      End
      Begin MSDataListLib.DataCombo cboSubGrupo 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   1170
         Width           =   2370
         _ExtentX        =   4180
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Filtrado por "
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   900
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Grupo "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   2850
      TabIndex        =   8
      Top             =   30
      Width           =   2745
      Begin VB.CheckBox chkGrupo 
         Caption         =   "Todos"
         Height          =   285
         Left            =   180
         TabIndex        =   2
         Top             =   330
         Value           =   1  'Checked
         Width           =   855
      End
      Begin MSDataListLib.DataCombo cboGrupo 
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Top             =   1170
         Width           =   2490
         _ExtentX        =   4392
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Filtrado por "
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   900
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Tipo de Producto "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   60
      TabIndex        =   6
      Top             =   30
      Width           =   2745
      Begin VB.CheckBox chkTipoProducto 
         Caption         =   "Todos"
         Height          =   285
         Left            =   150
         TabIndex        =   0
         Top             =   330
         Value           =   1  'Checked
         Width           =   855
      End
      Begin MSDataListLib.DataCombo cboTipoProducto 
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Top             =   1170
         Width           =   2490
         _ExtentX        =   4392
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Filtrado por "
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   870
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmFiltroProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                    ByVal hwnd As Long, _
                    ByVal lpOperation As String, _
                    ByVal lpFile As String, _
                    ByVal lpParameters As String, _
                    ByVal lpDirectory As String, _
                    ByVal nShowCmd As Long) As Long

Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWMINIMIZED As Long = 2

Option Explicit
Dim RsGrupo As Recordset
Dim RsSubGrupo As Recordset
Dim RsTipoProducto As Recordset
Dim rsCanalVenta As Recordset
Dim sCanal As String
Dim xCriterio As String
Dim xpatio As String

Sub LlenaCombos()
    With cboSubGrupo
         Isql = "Select * from vSubGrupo order by Descripcion"
         Set RsSubGrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubGrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    
    
    With cboGrupo
         Isql = "Select * from vGrupo order by Descripcion"
         Set RsGrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsGrupo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboTipoProducto
         Isql = "Select * from vTipoProducto order by Descripcion"
         Set RsTipoProducto = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoProducto
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
        With cboCanalVenta
         Isql = "select codigo, Descripcion from vTipoPedido Where lActivo = 1 order by descripcion"
         Set rsCanalVenta = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = rsCanalVenta
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
End Sub

Private Sub cboGrupo_Click(Area As Integer)
   chkGrupo.Value = False
End Sub

Private Sub cboGrupo_Change()
    cboSubGrupo.BoundText = ""
    With cboSubGrupo
         Isql = "Select * from vSubGrupo " & IIf(cboGrupo.BoundText = "", "", "where tGrupo = '" & cboGrupo.BoundText & "'") & " order by Descripcion "
         Set RsSubGrupo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubGrupo
    End With

End Sub

Private Sub cboSubGrupo_Click(Area As Integer)
   chkSubGrupo.Value = False
End Sub

Private Sub cboTipoProducto_Click(Area As Integer)
   chkTipoProducto.Value = False
End Sub

Private Sub chkGrupo_Click()
   If chkGrupo.Value = 1 Then
      cboGrupo.BoundText = ""
   End If
End Sub

Private Sub chkPatio_Click()
 If chkPatio.Value = 1 Then
      xpatio = " and lActivaPatioComida = 1 "
    Else
      xpatio = " "
   End If
End Sub

Private Sub chkSubGrupo_Click()
   If chkSubGrupo.Value = 1 Then
      cboSubGrupo.BoundText = ""
   End If
End Sub

Private Sub chkTipoProducto_Click()
   If chkTipoProducto.Value = 1 Then
      cboTipoProducto.BoundText = ""
   End If
End Sub

Private Sub chkVineta_Click()
    If chkVineta.Value = 1 Then
            cmdOpcion(4).Enabled = True
            Frame4.Enabled = False
            Frame5.Enabled = False
            Me.frameCanal.Enabled = False
            cmdOpcion(0).Enabled = False
            cmdOpcion(1).Enabled = False
            cmdOpcion(2).Enabled = False
    Else
            cmdOpcion(4).Enabled = False
            Frame4.Enabled = True
            Frame5.Enabled = True
            Me.frameCanal.Enabled = True
            cmdOpcion(0).Enabled = True
            cmdOpcion(1).Enabled = True
            cmdOpcion(2).Enabled = True

    End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   wEnter = True
   sTipoCanalReporte = ""
   sTipoCanalNombreReporte = ""
    
   If Index = 3 Then
               wEnter = False
               sTipo = "Salir"
                   Screen.MousePointer = vbDefault
                      Unload Me
               Exit Sub
               
   End If
               
   If Index = 4 Then
                Dim fparametros As String
                Dim pParametros As String
                Dim rstRegistros As ADODB.Recordset
                If Calcular("select count(*) codigo from timpresora where tcaja='" & sCaja & "' and isnull(lvineta,0)=1", Cn) = 0 Then
                    MsgBox "No existe Impresora de Viñetas para esta Caja", vbInformation, sMensaje
                         Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                If FileExists(App.Path & "\GeneracionEtiqueta\Generacionetiqueta.exe") = False Then
                    MsgBox "No tiene el Módulo de Impresión de Viñetas " & vbCrLf & "Comuníquese con Infhotel Servicios Informáticos", vbInformation, sMensaje
                         Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                If chkTipoProducto.Value = 0 Then
                    If cboTipoProducto.Text = "" Then
                         Screen.MousePointer = vbDefault
                        cboTipoProducto.SetFocus
                    End If
                End If
                If chkGrupo.Value = 0 Then
                    If cboGrupo.Text = "" Then
                         Screen.MousePointer = vbDefault
                        cboGrupo.SetFocus
                    End If
                End If
                If chkSubGrupo.Value = 0 Then
                    If cboSubGrupo.Text = "" Then
                         Screen.MousePointer = vbDefault
                        cboSubGrupo.SetFocus
                    End If
                End If
                
               'impresion viñetas
             '  MsgBox "viñetas"
                    ' Shell App.Path & "\GeneracionEtiqueta\GeneracionEtiqueta.exe" & " " & "0" + txtCodigo.Text + sCaja, vbNormalFocus
               If chkTipoProducto.Value = 1 Then
                    fparametros = "XX"
                    pParametros = "''"
                Else
                    fparametros = cboTipoProducto.BoundText
                    pParametros = cboTipoProducto.BoundText
                End If
                
               If chkGrupo.Value = 1 Then
                    fparametros = fparametros + "XX"
                    pParametros = pParametros + ",''"
                Else
                    fparametros = fparametros + cboGrupo.BoundText
                    pParametros = pParametros + "," + cboGrupo.BoundText
                End If
                
               If chkSubGrupo.Value = 1 Then
                    fparametros = fparametros + "XXXX"
                    pParametros = pParametros + ",''"
                Else
                    fparametros = fparametros + cboSubGrupo.BoundText
                    pParametros = pParametros + "," + cboSubGrupo.BoundText
                End If
                 
               Set rstRegistros = Lib.OpenRecordset("SP_GENERA_ETIQUETA " + pParametros, Cn)
               If rstRegistros.RecordCount = 0 Then
                    MsgBox "No existen registros a Imprimir", vbInformation, sMensaje
                     Screen.MousePointer = vbDefault
                    Exit Sub
               End If
                
               'Shell App.Path & "\GeneracionEtiqueta\GeneracionEtiqueta.exe" & " " & "1" + fparametros + sCaja, vbNormalFocus
               ShellExecute Me.hwnd, "open", App.Path & "\GeneracionEtiqueta\GeneracionEtiqueta.exe", "1" + fparametros + sCaja, "C:\", SW_SHOWNORMAL
        
               Sleep 3000
               Screen.MousePointer = vbDefault
               Unload Me
               Exit Sub

   End If
               

   If optTipo(2).Value = False Then
        If cboCanalVenta.Text = "" Then
                MsgBox "Debe seleccionar un Canal de Venta", vbInformation, sMensaje
                   Screen.MousePointer = vbDefault
                Exit Sub
        End If
   
   End If
 
         If optTipo(1).Value Then
           sDescrip = "Combo"
           sTipoCanalReporte = cboCanalVenta.BoundText
           sTipoCanalNombreReporte = cboCanalVenta.Text
         Else
         If optTipo(2).Value Then
           sDescrip = "Precios"
           sTipoCanalReporte = ""
           sTipoCanalNombreReporte = ""
         Else
           sDescrip = ""
           sTipoCanalReporte = cboCanalVenta.BoundText
           sTipoCanalNombreReporte = cboCanalVenta.Text
        End If
   End If
   
   
   
   Select Case Index
          Case Is = 0 ' Pantalla
               sTipo = "Pantalla"
               Genera
          Case Is = 1 ' Impresora
               sTipo = "Impresora"
               Genera
          Case Is = 2 ' XLS
               sTipo = "Excel"
               Genera
          Case Is = 3 ' Salir
               Screen.MousePointer = vbDefault
               wEnter = False
               sTipo = "Salir"
   End Select
   Unload Me
End Sub

Private Sub Form_Load()
   Centrar Me
   LlenaCombos
   xCriterio = ""
'   sCanal = Calcular("select tResumido as Codigo from ttabla where ttabla='TipoPedido' and tCodigo='01'", Cn)
'   optTipo(0).Caption = "Productos sin Combinaciones (" & sCanal & ")"
'   optTipo(1).Caption = "Productos con Combinaciones (" & sCanal & ")"
   
End Sub

Public Sub Genera()
      
    
    If chkTipoProducto.Value And chkGrupo.Value And chkSubGrupo And optEstado(0).Value Then
        If chkPatio.Value = 1 Then
         xpatio = " lActivaPatioComida = 1 "
         Else
           xpatio = " "
        End If
    
       xCriterio = ""
       sCodigo = " " + xpatio
       Exit Sub
    End If
    
     If optTipo(1).Value Then
       If chkTipoProducto = False Then
        If sDescrip = "" Then
             xCriterio = " tTipoProducto =''" & cboTipoProducto.BoundText & "''"
        ElseIf sDescrip = "Combo" Then
             xCriterio = " tTipoProducto =''" & cboTipoProducto.BoundText & "''"
        
            Else
             xCriterio = " tTipoProducto ='" & cboTipoProducto.BoundText & "'"
        End If
         
       End If
       
       If chkGrupo.Value = False Then
          If xCriterio = "" Then
             xCriterio = " tGrupo =''" & cboGrupo.BoundText & "''"
          Else
                If sDescrip = "" Then
                    xCriterio = xCriterio & " and tGrupo =''" & cboGrupo.BoundText & "''"
                ElseIf sDescrip = "Combo" Then
                    xCriterio = " tTipoProducto =''" & cboGrupo.BoundText & "''"
                Else
                    xCriterio = xCriterio & " and tGrupo ='" & cboGrupo.BoundText & "'"
                End If
             
          End If
       End If
       
       If chkSubGrupo.Value = False Then
          If xCriterio = "" Then
           If sDescrip = "Combo" Then
             xCriterio = " tSubGrupo =''" & cboSubGrupo.BoundText & "''"
          ElseIf sDescrip = "Combo" Then
             xCriterio = " tTipoProducto =''" & cboSubGrupo.BoundText & "''"
            Else
             xCriterio = " tSubGrupo ='" & cboSubGrupo.BoundText & "'"
             End If
          Else
             If sDescrip = "" Then
                xCriterio = xCriterio & " and tSubGrupo ='" & cboSubGrupo.BoundText & "'"
             Else
                xCriterio = xCriterio & " and tSubGrupo =''" & cboSubGrupo.BoundText & "''"
             End If
             
          End If
       End If
       
    ElseIf optTipo(2).Value Then
       If chkTipoProducto = False Then
          xCriterio = " tTipoProducto ='" & cboTipoProducto.BoundText & "'"
       End If
       
       If chkGrupo.Value = False Then
          If xCriterio = "" Then
             xCriterio = " tGrupo ='" & cboGrupo.BoundText & "'"
          Else
            If sDescrip = "" Then
                xCriterio = xCriterio & " and tGrupo =''" & cboGrupo.BoundText & "''"
            Else
                xCriterio = xCriterio & " and tGrupo ='" & cboGrupo.BoundText & "'"
            End If
            
             
          End If
       End If
       
       If chkSubGrupo.Value = False Then
          If xCriterio = "" Then
             xCriterio = " tSubGrupo ='" & cboSubGrupo.BoundText & "'"
          Else
            If sDescrip = "" Then
                xCriterio = xCriterio & " and tSubGrupo =''" & cboSubGrupo.BoundText & "''"
            Else
                xCriterio = xCriterio & " and tSubGrupo ='" & cboSubGrupo.BoundText & "'"
            End If
             
          End If
       End If
       
    Else
       If chkTipoProducto = False Then
          If sDescrip = "" Then
            xCriterio = "tTipoProducto =''" & cboTipoProducto.BoundText & "''"
          Else
            xCriterio = "tTipoProducto ='" & cboTipoProducto.BoundText & "'"
          End If
          
       End If
       
       If chkGrupo.Value = False Then
          If xCriterio = "" Then
          If sDescrip = "" Then
             xCriterio = "tGrupo =''" & cboGrupo.BoundText & "''"
             Else
               xCriterio = "tGrupo ='" & cboGrupo.BoundText & "'"
             End If
          Else
            If sDescrip = "" Then
                xCriterio = xCriterio & " and tGrupo =''" & cboGrupo.BoundText & "''"
            Else
                xCriterio = xCriterio & " and tGrupo ='" & cboGrupo.BoundText & "'"
            End If
             
          End If
       End If
       
       If chkSubGrupo.Value = False Then
          If xCriterio = "" Then
             xCriterio = "tSubGrupo =''" & cboSubGrupo.BoundText & "''"
          Else
             If sDescrip = "" Then
                xCriterio = xCriterio & " and tSubGrupo =''" & cboSubGrupo.BoundText & "''"
             Else
                xCriterio = xCriterio & " and tSubGrupo ='" & cboSubGrupo.BoundText & "'"
             End If
             
          End If
       End If
    End If
    
    If optEstado(1).Value Then
        If xCriterio = "" Then
           xCriterio = "lActivo=1"
        Else
          xCriterio = xCriterio & " and lActivo=1"
        End If
    End If
    
    If optEstado(2).Value Then
        If xCriterio = "" Then
           xCriterio = "lActivo=0"
        Else
          xCriterio = xCriterio & " and lActivo=0"
        End If

    End If
    
    sCodigo = xCriterio + xpatio
End Sub

Private Sub optTipo_Click(Index As Integer)
        Select Case Index
            Case 0
                    If optTipo(0).Value = True Then
                        frameCanal.Enabled = True
                    Else
                        frameCanal.Enabled = False
                    End If
            Case 1
                    If optTipo(1).Value = True Then
                        frameCanal.Enabled = True
                    Else
                        frameCanal.Enabled = False
                    End If
            Case 2
                
                    If optTipo(2).Value = True Then
                        frameCanal.Enabled = False
                    
                    End If
        End Select
End Sub
