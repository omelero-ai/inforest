VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmParametroCheffControl 
   Caption         =   "Parámetros del Cheff Control"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5640
   Icon            =   "frmParametroCheffControl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   0
      TabIndex        =   12
      Top             =   2115
      Width           =   5595
   End
   Begin VB.TextBox txtDias 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2790
      TabIndex        =   11
      Text            =   "1.00"
      Top             =   1755
      Width           =   825
   End
   Begin VB.CheckBox chkCanal 
      Caption         =   "Canal de Venta"
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
      Index           =   5
      Left            =   90
      TabIndex        =   7
      Top             =   1800
      Width           =   2265
   End
   Begin VB.CheckBox chkCanal 
      Caption         =   "Canal de Venta"
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
      Index           =   4
      Left            =   90
      TabIndex        =   6
      Top             =   1473
      Width           =   2265
   End
   Begin VB.CheckBox chkCanal 
      Caption         =   "Canal de Venta"
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
      Index           =   3
      Left            =   90
      TabIndex        =   5
      Top             =   1147
      Width           =   2265
   End
   Begin VB.CheckBox chkCanal 
      Caption         =   "Canal de Venta"
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
      Index           =   2
      Left            =   90
      TabIndex        =   4
      Top             =   821
      Width           =   2265
   End
   Begin VB.CheckBox chkCanal 
      Caption         =   "Canal de Venta"
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
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   495
      Width           =   2265
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Grabar"
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
      Left            =   4425
      Picture         =   "frmParametroCheffControl.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2250
      Width           =   1170
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Cancelar"
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
      Left            =   3195
      Picture         =   "frmParametroCheffControl.frx":0DFC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2250
      Width           =   1170
   End
   Begin MSDataListLib.DataCombo cboEstadoPedido 
      Height          =   315
      Left            =   2790
      TabIndex        =   8
      Top             =   900
      Width           =   2835
      _ExtentX        =   5001
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
      Caption         =   "Cantidad de días a Revisar a los pedidos no atendidos"
      Height          =   420
      Left            =   2790
      TabIndex        =   10
      Top             =   1305
      Width           =   2670
   End
   Begin VB.Label Label1 
      Caption         =   "Escoger el Estado del Pedido que activara el pedido en el Cheff Control"
      Height          =   420
      Left            =   2835
      TabIndex        =   9
      Top             =   450
      Width           =   2760
   End
   Begin VB.Label Label5 
      Caption         =   "Canales de Venta"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   135
      Width           =   1635
   End
End
Attribute VB_Name = "frmParametroCheffControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsEstadoPedido As Recordset
Dim RsParametro As Recordset

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 0 Then
      Isql = "Update TPARAMETRO Set " & _
             "lCheffControlCanal1 =" & IIf(chkCanal(1).Value, 1, 0) & ", " & _
             "lCheffControlCanal2 =" & IIf(chkCanal(2).Value, 1, 0) & ", " & _
             "lCheffControlCanal3 =" & IIf(chkCanal(3).Value, 1, 0) & ", " & _
             "lCheffControlCanal4 =" & IIf(chkCanal(4).Value, 1, 0) & ", " & _
             "lCheffControlCanal5 =" & IIf(chkCanal(5).Value, 1, 0) & ", " & _
             "tCheffControlEstadoPedido = '" & cboEstadoPedido.BoundText & "', " & _
             "nDiasCheffControl = " & Val(txtDias.Text)
      Cn.Execute Isql
      MsgBox "Datos Actualizados", vbInformation, sMensaje
      Unload Me
   Else
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Centrar Me
   With cboEstadoPedido
        Isql = "Select * from vEstadoPedido where lActivo = 1 order by Codigo"
        Set RsEstadoPedido = Lib.OpenRecordset(Isql, Cn)
        Set .RowSource = RsEstadoPedido
            .DataField = "Descripcion"
            .ListField = "Descripcion"
            .BoundColumn = "Codigo"
   End With
   
   Isql = "select * from TPARAMETRO"
   Set RsParametro = Lib.OpenRecordset(Isql, Cn)
   
   If Not RsParametro.EOF Then
      RsParametro.MoveFirst
      
      If IsNull(RsParametro!tBoton1) Or RsParametro!tBoton1 = "" Then
         chkCanal(1).Caption = "(no definido)"
         chkCanal(1).Enabled = False
         chkCanal(1).Value = 0
      Else
         chkCanal(1).Caption = RsParametro!tBoton1
         chkCanal(1).Value = IIf(IsNull(RsParametro!lCheffControlCanal1), 0, IIf(RsParametro!lCheffControlCanal1 = True, 1, 0))
      End If
      If IsNull(RsParametro!tBoton2) Or RsParametro!tBoton2 = "" Then
         chkCanal(2).Caption = "(no definido)"
         chkCanal(2).Enabled = False
         chkCanal(3).Value = 0
      Else
         chkCanal(2).Caption = RsParametro!tBoton2
         chkCanal(2).Value = IIf(IsNull(RsParametro!lCheffControlCanal2), 0, IIf(RsParametro!lCheffControlCanal2 = True, 1, 0))
      End If
      If IsNull(RsParametro!tBoton3) Or RsParametro!tBoton3 = "" Then
         chkCanal(3).Caption = "(no definido)"
         chkCanal(3).Enabled = False
         chkCanal(3).Value = 0
      Else
         chkCanal(3).Caption = RsParametro!tBoton3
         chkCanal(3).Value = IIf(IsNull(RsParametro!lCheffControlCanal3), 0, IIf(RsParametro!lCheffControlCanal3 = True, 1, 0))
      End If
      If IsNull(RsParametro!tBoton4) Or RsParametro!tBoton4 = "" Then
         chkCanal(4).Caption = "(no definido)"
         chkCanal(4).Enabled = False
         chkCanal(4).Value = 0
      Else
         chkCanal(4).Caption = RsParametro!tBoton4
         chkCanal(4).Value = IIf(IsNull(RsParametro!lCheffControlCanal4), 0, IIf(RsParametro!lCheffControlCanal4 = True, 1, 0))
      End If
      If IsNull(RsParametro!tBoton5) Or RsParametro!tBoton5 = "" Then
         chkCanal(5).Caption = "(no definido)"
         chkCanal(5).Enabled = False
         chkCanal(5).Value = 0
      Else
         chkCanal(5).Caption = RsParametro!tBoton5
         chkCanal(5).Value = IIf(IsNull(RsParametro!lCheffControlCanal5), 0, IIf(RsParametro!lCheffControlCanal5 = True, 1, 0))
      End If
      cboEstadoPedido.BoundText = IIf(IsNull(RsParametro!tCheffControlEstadoPedido), "", Trim(RsParametro!tCheffControlEstadoPedido))
      txtDias.Text = Format(RsParametro!nDiasCheffControl, "##0.00")
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsParametro = Nothing
   Set frmParametroCheffControl = Nothing
End Sub

Private Sub txtDias_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   Else
      Numerico KeyAscii, txtDias
   End If
End Sub

Private Sub txtDias_LostFocus()
   txtDias.Text = Format(txtDias.Text, "###,##0.00")
End Sub
