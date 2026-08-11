VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmFiltroPropiedad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emite de Propiedades"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "frmFiltroPropiedad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
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
      Left            =   4770
      TabIndex        =   19
      Top             =   1800
      Width           =   2895
      Begin VB.OptionButton optEstado 
         Caption         =   "No Activos"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   22
         Top             =   855
         Width           =   2310
      End
      Begin VB.OptionButton optEstado 
         Caption         =   "Activos"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   21
         Top             =   585
         Width           =   2310
      End
      Begin VB.OptionButton optEstado 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   20
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
      Left            =   45
      TabIndex        =   16
      Top             =   1800
      Width           =   4695
      Begin VB.OptionButton optTipo 
         Caption         =   "Visualizacion de Propiedades"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   18
         Top             =   585
         Width           =   4290
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Productos"
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   17
         Top             =   315
         Value           =   -1  'True
         Width           =   4245
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
      Index           =   2
      Left            =   4770
      Picture         =   "frmFiltroPropiedad.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3090
      Visible         =   0   'False
      Width           =   1455
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
      Left            =   5160
      TabIndex        =   13
      Top             =   30
      Width           =   2505
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
         Width           =   2250
         _ExtentX        =   3969
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
         TabIndex        =   14
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
      Left            =   2610
      TabIndex        =   11
      Top             =   30
      Width           =   2505
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
         Width           =   2250
         _ExtentX        =   3969
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
         TabIndex        =   12
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
      TabIndex        =   9
      Top             =   30
      Width           =   2505
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
         Width           =   2250
         _ExtentX        =   3969
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
         TabIndex        =   10
         Top             =   870
         Width           =   825
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
      Left            =   1860
      Picture         =   "frmFiltroPropiedad.frx":082E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3090
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
      Picture         =   "frmFiltroPropiedad.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3090
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
      Left            =   6225
      Picture         =   "frmFiltroPropiedad.frx":1292
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3090
      Width           =   1455
   End
End
Attribute VB_Name = "frmFiltroPropiedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsGrupo As Recordset
Dim RsSubGrupo As Recordset
Dim RsTipoProducto As Recordset
Dim sCanal As String
Dim xCriterio As String

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

Private Sub cmdOpcion_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   wEnter = True
   If optTipo(1).Value Then
      sDescrip = "Propiedades"
   Else
    If optTipo(0).Value Then
      sDescrip = "Productos"
    Else
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
   sCanal = Calcular("select tResumido as Codigo from ttabla where ttabla='TipoPedido' and tCodigo='01'", Cn)
'   optTipo(0).Caption = "Productos sin Combinaciones (" & sCanal & ")"
'   optTipo(1).Caption = "Productos con Combinaciones (" & sCanal & ")"
   
End Sub

Public Sub Genera()
    If chkTipoProducto.Value And chkGrupo.Value And chkSubGrupo And optEstado(0).Value Then
       xCriterio = ""
       sCodigo = ""
       Exit Sub
    End If
    
'    If optTipo(1).Value Then
'       If chkTipoProducto = False Then
'          xCriterio = "vProducto.tTipoProducto ='" & cboTipoProducto.BoundText & "'"
'       End If
'
'       If chkGrupo.Value = False Then
'          If xCriterio = "" Then
'             xCriterio = "vProducto.tGrupo ='" & cboGrupo.BoundText & "'"
'          Else
'             xCriterio = xCriterio & " and vProducto.tGrupo ='" & cboGrupo.BoundText & "'"
'          End If
'       End If
'
'       If chkSubGrupo.Value = False Then
'          If xCriterio = "" Then
'             xCriterio = "vProducto.tSubGrupo ='" & cboSubGrupo.BoundText & "'"
'          Else
'             xCriterio = xCriterio & " and vProducto.tSubGrupo ='" & cboSubGrupo.BoundText & "'"
'          End If
'       End If
'    Else
       If chkTipoProducto = False Then
          xCriterio = "tTipoProducto ='" & cboTipoProducto.BoundText & "'"
       End If
       
       If chkGrupo.Value = False Then
          If xCriterio = "" Then
             xCriterio = "tGrupo ='" & cboGrupo.BoundText & "'"
          Else
             xCriterio = xCriterio & " and tGrupo ='" & cboGrupo.BoundText & "'"
          End If
       End If
       
       If chkSubGrupo.Value = False Then
          If xCriterio = "" Then
             xCriterio = "tSubGrupo ='" & cboSubGrupo.BoundText & "'"
          Else
             xCriterio = xCriterio & " and tSubGrupo ='" & cboSubGrupo.BoundText & "'"
          End If
       End If
'    End If
    
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
    
    sCodigo = xCriterio
End Sub

