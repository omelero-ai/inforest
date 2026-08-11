VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRepCtaCte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estados de Cuentas Corrientes"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   Icon            =   "frmRepCtaCte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
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
      Left            =   5355
      Picture         =   "frmRepCtaCte.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4590
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
      Index           =   2
      Left            =   6825
      Picture         =   "frmRepCtaCte.frx":082E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4590
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
      Left            =   3900
      Picture         =   "frmRepCtaCte.frx":0920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4590
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
      Left            =   2445
      Picture         =   "frmRepCtaCte.frx":0E52
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4590
      Width           =   1455
   End
   Begin VB.Frame Frame 
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
      Height          =   4500
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8265
      Begin VB.CheckBox chkSubTipo 
         Caption         =   "Todos los SubTipos"
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
         Left            =   6045
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2085
      End
      Begin VB.CheckBox chkTipo 
         Caption         =   "Todos los Tipos"
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
         Left            =   6045
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1255
         Value           =   1  'Checked
         Width           =   1995
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
         Left            =   6075
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   830
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.ComboBox cboEstado 
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
         Height          =   315
         ItemData        =   "frmRepCtaCte.frx":1384
         Left            =   2520
         List            =   "frmRepCtaCte.frx":138E
         TabIndex        =   20
         Top             =   816
         Width           =   3345
      End
      Begin VB.Frame Frame1 
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
         Height          =   1320
         Left            =   2520
         TabIndex        =   15
         Top             =   3060
         Width           =   3345
         Begin VB.OptionButton optOpcion 
            Caption         =   "Consolidado"
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
            Left            =   450
            TabIndex        =   18
            Top             =   315
            Width           =   1545
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
            Height          =   240
            Index           =   2
            Left            =   450
            TabIndex        =   17
            Top             =   900
            Value           =   -1  'True
            Width           =   1545
         End
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
            Height          =   240
            Index           =   1
            Left            =   450
            TabIndex        =   16
            Top             =   607
            Width           =   1545
         End
      End
      Begin VB.TextBox txtCliente 
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
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   390
         Width           =   2625
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
         Height          =   330
         Left            =   5235
         Picture         =   "frmRepCtaCte.frx":13AD
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   382
         Width           =   630
      End
      Begin MSMAPI.MAPIMessages MAPIM 
         Left            =   6570
         Top             =   3015
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         AddressEditFieldCount=   1
         AddressModifiable=   0   'False
         AddressResolveUI=   0   'False
         FetchSorted     =   0   'False
         FetchUnreadOnly =   0   'False
      End
      Begin MSMAPI.MAPISession MAPIS 
         Left            =   6570
         Top             =   3645
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DownloadMail    =   0   'False
         LogonUI         =   0   'False
         NewSession      =   0   'False
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
         Left            =   6075
         TabIndex        =   8
         Top             =   405
         Value           =   1  'Checked
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   2580
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
         Format          =   63963137
         CurrentDate     =   37541.9993055556
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   375
         Left            =   2520
         TabIndex        =   0
         Top             =   2094
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
         Format          =   63963137
         CurrentDate     =   37539.2083333333
      End
      Begin MSComDlg.CommonDialog cmdSave 
         Left            =   6615
         Top             =   2385
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo cboTipo 
         Height          =   315
         Left            =   2520
         TabIndex        =   27
         Top             =   1242
         Width           =   3345
         _ExtentX        =   5900
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
      Begin MSDataListLib.DataCombo cboSubTipo 
         Height          =   315
         Left            =   2520
         TabIndex        =   28
         Top             =   1668
         Width           =   3345
         _ExtentX        =   5900
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
      Begin MSComCtl2.DTPicker dtpHoraIni 
         Height          =   375
         Left            =   4515
         TabIndex        =   1
         Top             =   2094
         Width           =   1350
         _ExtentX        =   2381
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
         CurrentDate     =   37541
      End
      Begin MSComCtl2.DTPicker dtpHoraFin 
         Height          =   375
         Left            =   4500
         TabIndex        =   3
         Top             =   2580
         Width           =   1350
         _ExtentX        =   2381
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
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SubTipo Cuenta Corriente :"
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
         Left            =   135
         TabIndex        =   26
         Top             =   1725
         Width           =   2325
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Cuenta Corriente:"
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
         Left            =   525
         TabIndex        =   24
         Top             =   1305
         Width           =   1935
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Pedido :"
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
         Left            =   1095
         TabIndex        =   22
         Top             =   875
         Width           =   1365
      End
      Begin VB.Label Label 
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
         Index           =   0
         Left            =   1740
         TabIndex        =   14
         Top             =   450
         Width           =   720
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
         Left            =   1335
         TabIndex        =   13
         Top             =   2665
         Width           =   1125
      End
      Begin VB.Label Label2 
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
         Left            =   1230
         TabIndex        =   12
         Top             =   2180
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "E-mail"
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
      Left            =   5085
      Picture         =   "frmRepCtaCte.frx":14AF
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   1455
   End
End
Attribute VB_Name = "frmRepCtaCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsReporte As Recordset
Dim RsCliente As Recordset
Dim RsTipo As Recordset
Dim RsSubTipo As Recordset
Dim ReportC As New dsrCtaCteC
Dim ReportR As New dsrCtaCteR
Dim ReportD As New dsrCtaCteD
Dim sCriterio As String
Dim sEstadoPedido As String
Dim sTitulo As String
Dim fInicio As Date
Dim fFinal As Date

Sub LlenaCombos()
    With cboTipo
         Isql = "Select * from vTipoCtaCte where lActivo=1"
         Set RsTipo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipo
         .DataField = "Descripcion"
         .ListField = "Descripcion"
         .BoundColumn = "Codigo"
    End With
    
    With cboSubTipo
         Isql = "Select * from vSubTipoCtaCte where lActivo=1"
         Set RsSubTipo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubTipo
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
End Sub

Private Sub cboTipo_Change()
    cboSubTipo.BoundText = ""
    With cboSubTipo
         Isql = "Select * from vSubTipoCtaCte " & IIf(cboTipo.BoundText = "", "", "where tTipoCtaCte = '" & cboTipo.BoundText & "'") & " order by Descripcion "
         Set RsSubTipo = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubTipo
    End With
End Sub

Private Sub chkEstado_Click()
   If chkEstado.Value = 1 Then
      cboEstado.Enabled = False
      cboEstado.Text = ""
   Else
      cboEstado.Enabled = True
   End If
End Sub

Private Sub chkSubTipo_Click()
   If chkSubTipo.Value = 1 Then
      cboSubTipo.Enabled = False
      cboSubTipo.Text = ""
   Else
      cboSubTipo.Enabled = True
   End If
End Sub

Private Sub chkTipo_Click()
   If chkTipo.Value = 1 Then
      cboTipo.Enabled = False
      cboTipo.Text = ""
   Else
      cboTipo.Enabled = True
   End If
End Sub

Private Sub cmdBusca_Click()
   Dim xCriterio As String
   Isql = "Select tCodigoCliente as Codigo, tIdentidad, tApecom as Descripcion from TCOMPANIA order by tApeCom"
   
    Isql = "Select Codigo as Codigo, Identidad, Descripcion as Descripcion from vCOMPANIA order by descripcion"
 
   frmBusca.cboCriterio.Enabled = True
   frmBusca.nPredeterm = 2
   Call ConfGrilla(3, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                          "Identificador", 2, "tIdentidad", 1500, 0, 0, "", _
                                          "Razón Comercial", 2, "Descripcion", 4500, 0, 0, "")
   frmBusca.Show vbModal
   If Not wEnter Then
      Exit Sub
   End If
   sCliente = sCodigo
   txtCliente.Text = sDescrip
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

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 2 Then
      Unload Me
      Exit Sub
   End If
   sEstadoPedido = ""
  ' sCriterio = "dbo.MPEDIDO.fRegistro >= '" & Format(dtpFecIni.value, "yyyy/mm/dd") & " " & Format(dtpHoraIni.value, "HH:mm") & "' and dbo.MPEDIDO.fRegistro <= '" & Format(dtpFecFin.value, "yyyy/mm/dd") & " " & Format(dtpHoraFin.value, "HH:mm") & "'"
   sTitulo = ""
   sCriterio = ""
   If chkCliente.Value = 0 Then
      If sCliente = "" Then
         MsgBox "Debe escoger un Cliente", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & " tClienteCtaCte ='" & sCliente & "'"
   End If
       
   If chkEstado.Value = 0 Then
      If cboEstado.Text = "" Then
         MsgBox "Debe escoger un Estado de Pedido", vbCritical, sMensaje
         Exit Sub
      End If
      If cboEstado.Text = "Facturados" Then
         sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & " tEstadoPedido='02'"
         sEstadoPedido = "02"
         sTitulo = "Pedidos Facturados"
      Else
         sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & " tEstadoPedido='04'"
         sEstadoPedido = "04"
         sTitulo = "Pedidos No Facturados"
      End If
   Else
      sTitulo = "Todos los Pedidos"
   End If
   
   If chkTipo.Value = 0 Then
      If cboTipo.BoundText = "" Then
         MsgBox "Debe escoger un Tipo", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & " tTipoCtaCte ='" & cboTipo.BoundText & "'"
   End If
   
   If chkSubTipo.Value = 0 Then
      If cboSubTipo.BoundText = "" Then
         MsgBox "Debe escoger un SubTipo", vbCritical, sMensaje
         Exit Sub
      End If
      sCriterio = sCriterio & IIf(sCriterio = "", "", " and ") & " tSubTipoCtaCte ='" & cboSubTipo.BoundText & "'"
   End If
   If chkTipo.Value = 0 Then
      sTitulo = sTitulo & ", Tipo " & cboTipo.Text
   End If
   If chkSubTipo.Value = 0 Then
      sTitulo = sTitulo & ", SubTipo " & cboSubTipo.Text
   End If
   sTitulo = sTitulo & " Del " & Format(dtpFecIni.Value, "dd/mmm/yyyy") & " al " & Format(dtpFecFin.Value, "dd/mmm/yyyy")
       
   Select Case Index
          Case Is = 0 ' Preview
               Genera
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               
               If optOpcion(0).Value = True Then
                  ReportC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReportC.PaperOrientation = crPortrait
               ElseIf optOpcion(1).Value = True Then
                  ReportR.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReportR.PaperOrientation = crPortrait
               Else
                  ReportD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReportD.PaperOrientation = crPortrait
               End If
               frmEmite.CRViewer.ViewReport
               frmEmite.Show vbModal
          
          Case Is = 1 ' Imprimir
               Genera
               If rsReporte.EOF = True Then
                   MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                   Screen.MousePointer = vbDefault
                   Exit Sub
               End If
               If optOpcion(0).Value = True Then
                  ReportC.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReportC.PaperOrientation = crPortrait
                  Screen.MousePointer = vbDefault
                  ReportC.PrintOut
               ElseIf optOpcion(1).Value = True Then
                  ReportR.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReportR.PaperOrientation = crPortrait
                  Screen.MousePointer = vbDefault
                  ReportR.PrintOut
               Else
                  ReportD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
                  ReportD.PaperOrientation = crPortrait
                  Screen.MousePointer = vbDefault
                  ReportD.PrintOut
               End If
          
          Case Is = 3 ' Mail
               Genera
               If rsReporte.EOF = True Then
                   MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                   Screen.MousePointer = vbDefault
                   Exit Sub
               End If
               Screen.MousePointer = vbDefault
               
               If chkCliente.Value = 0 Then
                  EnviaMail RsCliente!tEmail
               Else
                  RsCliente.MoveFirst
                  Do While Not RsCliente.EOF
                     rsReporte.Filter = "tClienteCtaCte = '" & RsCliente!codigo & "'"
                     If rsReporte.RecordCount <> 0 Then
                        If IsNull(RsCliente!tEmail) Or Len(Trim(RsCliente!tEmail)) = 0 Then
                        Else
                           EnviaMail RsCliente!tEmail
                        End If
                     End If
                     RsCliente.MoveNext
                  Loop
               End If
         
          Case Is = 4 ' Exportar
               Genera
               Screen.MousePointer = vbDefault
               If rsReporte.EOF = True Then
                  MsgBox "No hay Datos para Mostrar", vbExclamation, "Mensaje del Sistema"
                  Exit Sub
               End If
               
               If optOpcion(0).Value = True Then
                  ReportC.ExportOptions.FormatType = 21
                  ReportC.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  ReportC.ExportOptions.DiskFileName = cmdSave.FileName
                  ReportC.Export False
                  
               ElseIf optOpcion(1).Value Then
                  ReportR.ExportOptions.FormatType = 21
                  ReportR.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  ReportR.ExportOptions.DiskFileName = cmdSave.FileName
                  ReportR.Export False
               Else
                  ReportD.ExportOptions.FormatType = 21
                  ReportD.ExportOptions.DestinationType = 1
                  cmdSave.Filter = "Libro de Microsoft Excel|*.xls"
                  cmdSave.ShowSave
                  If cmdSave.FileName = "" Then
                     Exit Sub
                  End If
                  ReportD.ExportOptions.DiskFileName = cmdSave.FileName
                  ReportD.Export False
               End If
         
         Case Is = 2 ' Salir
              Unload Me
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
    dtpFecIni.Value = FechaServidor()
    dtpFecFin.Value = FechaServidor()
    cmdBusca.Enabled = False
    cboEstado.Enabled = False
    cboTipo.Enabled = False
    cboSubTipo.Enabled = False
    sCliente = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsReporte = Nothing
   Set RsCliente = Nothing
End Sub

Public Sub Genera()
   Screen.MousePointer = vbHourglass
   Dim oComando As clsComando
   Set oComando = New clsComando
   If Not oComando.CreateCmdSp("spRep_CtaCteN", Cn) Then
       Set oComando = Nothing
       Exit Sub
   End If
    fInicio = Format(dtpFecIni.Value, "yyyy/MM/dd") & " " & Format(dtpHoraIni.Value, "HH:mm")
    fFinal = Format(dtpFecFin.Value, "yyyy/MM/dd") & " " & Format(dtpHoraFin.Value, "HH:mm")
    
        oComando.CreateParameter "@flagTDetalle", adBoolean, adParamInput, 1, optOpcion(2).Value
        oComando.CreateParameter "@flagTResumido", adBoolean, adParamInput, 1, optOpcion(1).Value
        oComando.CreateParameter "@flagTConsolidado", adBoolean, adParamInput, 1, optOpcion(0).Value
        oComando.CreateParameter "@fInicio", adDBDate, adParamInput, 10, fInicio
        oComando.CreateParameter "@fFinal", adDBDate, adParamInput, 10, fFinal '
        oComando.CreateParameter "@Estado", adVarChar, adParamInput, 50, sEstadoPedido 'sEstadoPedido
        oComando.CreateParameter "@Cliente", adVarChar, adParamInput, 50, sCliente
        oComando.CreateParameter "@TipoCC", adVarChar, adParamInput, 50, cboTipo.BoundText
        oComando.CreateParameter "@SubTipoCC", adVarChar, adParamInput, 50, cboSubTipo.BoundText
   If Not oComando.GetParamOK Then
      Set oComando = Nothing
      Exit Sub
   End If
   Set rsReporte = oComando.GetSP()
   
   'If sCriterio <> "" Then
    'rsReporte.Filter = sCriterio
   'End If
   
   If optOpcion(0).Value = True Then

      ReportC.DiscardSavedData
      ReportC.Database.SetDataSource rsReporte
      ReportC.ReportTitle = sTitulo
      ReportC.Text15.SetText sRazonSocial
      ReportC.Text6.SetText localConectado
      frmEmite.CRViewer.DisplayGroupTree = False
      frmEmite.CRViewer.ReportSource = ReportC
      
   ElseIf optOpcion(1).Value = True Then

      ReportR.DiscardSavedData
      ReportR.Database.SetDataSource rsReporte
      ReportR.ReportTitle = sTitulo
      ReportR.Text15.SetText sRazonSocial
      ReportR.Text10.SetText localConectado
      frmEmite.CRViewer.DisplayGroupTree = False
      frmEmite.CRViewer.ReportSource = ReportR
   Else
        ReportD.Database.SetDataSource Nothing
      ReportD.DiscardSavedData
      ReportD.Database.SetDataSource rsReporte
      ReportD.ReportTitle = sTitulo
      ReportD.Text15.SetText sRazonSocial
      ReportD.Text20.SetText localConectado
      frmEmite.CRViewer.DisplayGroupTree = False
      frmEmite.CRViewer.ReportSource = ReportD
   End If
   
End Sub

Public Sub EnviaMail(Email As String)
    On Error GoTo MailError
    MAPIS.LogonUI = False
    MAPIS.SignOn
    If MAPIS.SessionID = 0 Then
       MAPIS.SessionID = -1
    End If
                   
    ReportD.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    ReportD.PaperSize = crPaperLetter
    ReportD.PaperOrientation = crPortrait
    ReportD.ExportOptions.DestinationType = 1
    ReportD.ExportOptions.DiskFileName = "C:\EstadoCta.doc"
    ReportD.ExportOptions.FormatType = crEFTRichText
    ReportD.Export False
          
    MAPIM.SessionID = MAPIS.SessionID
    MAPIM.Compose
    MAPIM.RecipAddress = Email
    MAPIM.RecipDisplayName = Email
    
    MAPIM.MsgSubject = "Estado de Cuenta Corriente"
    MAPIM.AttachmentPathName = "C:\EstadoCta.doc"
    MAPIM.send False
    MAPIS.SignOff
    DeleteFile "C:\EstadoCta.doc"
Exit Sub
MailError:
    MAPIS.SignOff
    MsgBox err.Number & " : " & err.Description
    DeleteFile "C:\EstadoCta.doc"
End Sub


