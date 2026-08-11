VERSION 5.00
Begin VB.Form frmProductoxProductoCantidad 
   Caption         =   "Detalle de Equivalencias de Producto"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7485
      Begin VB.TextBox txtCodigoProducto 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   315
         Width           =   1170
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
         Height          =   315
         Left            =   6705
         Picture         =   "frmProductoxProductoCantidad.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   690
         Width           =   585
      End
      Begin VB.TextBox txtProducto 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1185
         MaxLength       =   100
         TabIndex        =   0
         Text            =   " "
         Top             =   697
         Width           =   5400
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
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
         Left            =   1185
         MaxLength       =   100
         TabIndex        =   1
         Text            =   " "
         Top             =   1080
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   510
         TabIndex        =   9
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cantidad :"
         Height          =   195
         Left            =   375
         TabIndex        =   8
         Top             =   1125
         Width           =   720
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Equivalencia :"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   7
         Top             =   735
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdOpcionGrilla 
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
      Index           =   1
      Left            =   4950
      Picture         =   "frmProductoxProductoCantidad.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1620
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpcionGrilla 
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
      Index           =   2
      Left            =   6255
      Picture         =   "frmProductoxProductoCantidad.frx":0634
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1620
      Width           =   1215
   End
End
Attribute VB_Name = "frmProductoxProductoCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBusca_Click()
    Isql = "select Codigo, Descripcion, Grupo, SubGrupo, nPrecioVenta from vProducto where Codigo not in (select tSubProducto as Codigo from TPRODUCTOXPRODUCTO where tCodigoProducto='" & sVar1 & "') order by Descripcion"
    frmBusca.nPredeterm = 3
    Call ConfGrilla(5, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 800, 2, 0, "", _
                                           "Grupo", 2, "Grupo", 1300, 0, 0, "", _
                                           "SubGrupo", 2, "SubGrupo", 1400, 0, 0, "", _
                                           "Plato", 2, "Descripcion", 2900, 0, 0, "", _
                                           "Precio", 2, "nPrecioVenta", 900, 1, 0, "##,##0.00")
    frmBusca.Show vbModal
    If Not wEnter Then
       Exit Sub
    End If
    txtCodigoProducto.Text = sCodigo
    txtProducto.Text = sDescrip
    SendKeys "{Tab}"
End Sub

Private Sub cmdOpcionGrilla_Click(Index As Integer)
   If Index = 1 Then
      If txtCodigoProducto = "" Then MsgBox "Seleccione el producto", vbExclamation, sMensaje: cmdBusca.SetFocus: Exit Sub
         If frmProductoxProductoDetalle.lAgrega Then
            If Calcular("select tCodigoProducto as Codigo from TPRODUCTOXPRODUCTO where tCodigoProducto='" & sVar1 & "' and tSubProducto='" & txtCodigoProducto & "'", Cn) <> 0 Then
               MsgBox "El producto ya existe", vbExclamation, sMensaje
               Exit Sub
            End If
            
            Isql = "insert into TPRODUCTOXPRODUCTO " & _
                   "(tCodigoProducto, tSubProducto, nCantidad) " & _
                   "values ('" & sVar1 & "', '" & txtCodigoProducto.Text & "', " & Format(txtCantidad.Text, "####0.000") & ")"
            Cn.Execute Isql
            
            frmProductoxProductoDetalle.RsProductoDetalle.Sort = "tSubProducto ASC"
            frmProductoxProductoDetalle.RsProductoDetalle.Requery
            frmProductoxProductoDetalle.RsProductoDetalle.MoveLast
            MsgBox "Registro Agregado", vbInformation, sMensaje
            frmProductoxProductoDetalle.lAgrega = False
         Else
            nPos = frmProductoxProductoDetalle.RsProductoDetalle.AbsolutePosition
            Isql = "update TPRODUCTOXPRODUCTO set nCantidad=" & CDbl(txtCantidad.Text) & " where tCodigoProducto='" & sVar1 & "' and tSubProducto='" & txtCodigoProducto.Text & "'"
            Cn.Execute Isql
            frmProductoxProductoDetalle.RsProductoDetalle.Requery
            frmProductoxProductoDetalle.RsProductoDetalle.AbsolutePosition = nPos
            MsgBox "Registro Guardado", vbInformation, sMensaje
         End If
         Unload Me
   Else
      Unload Me
   End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me

    If frmProductoxProductoDetalle.lAgrega = True Then
       Blanquear Me
    Else
       Asignar
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtCantidad_GotFocus()
    Enfoque txtCantidad
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    TabNext KeyAscii
    Numerico KeyAscii, txtCantidad
End Sub

Private Sub txtCantidad_LostFocus()
   If txtCantidad.Text = "" Then
      txtCantidad.Text = "0"
   End If
   txtCantidad.Text = Format(txtCantidad.Text, "###,##0.000")
End Sub

Private Sub txtProducto_Change()
  If Len(Trim(txtProducto.Text)) < 2 Then
     Isql = "select Codigo, Descripcion, Grupo, SubGrupo, nPrecioVenta from vProducto where Codigo not in (select tSubProducto as Codigo from TPRODUCTOXPRODUCTO where tCodigoProducto='" & sVar1 & "') order by Descripcion"
     frmBusca.txtCriterio = Trim(txtProducto.Text)
     frmBusca.txtCriterio.SelStart = 2
     Call cmdBusca_Click
  End If
End Sub

Private Sub txtProducto_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      SendKeys "{Tab}"
   End If
End Sub
Public Sub Asignar()
    With frmProductoxProductoDetalle.RsProductoDetalle
         txtCodigoProducto = IIf(IsNull(!tSubProducto), "", !tSubProducto)
         txtProducto = IIf(IsNull(!producto), "", !producto)
         txtProducto.Enabled = False
         cmdBusca.Enabled = False
         txtCantidad = Format(IIf(IsNull(!nCantidad), 0, !nCantidad), "##,##0.000")
    End With
End Sub
