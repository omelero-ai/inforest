VERSION 5.00
Begin VB.Form frmCuentaPago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas Contables"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3450
   Icon            =   "frmCuentaPago.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   2175
      Picture         =   "frmCuentaPago.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
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
      Left            =   945
      Picture         =   "frmCuentaPago.frx":083C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1170
   End
   Begin VB.TextBox txtOtro 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1785
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1305
      Width           =   1635
   End
   Begin VB.TextBox txtDeposito 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1785
      MaxLength       =   15
      TabIndex        =   2
      Top             =   915
      Width           =   1635
   End
   Begin VB.TextBox txtEfectivoME 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1785
      MaxLength       =   15
      TabIndex        =   1
      Top             =   525
      Width           =   1635
   End
   Begin VB.TextBox txtEfectivoMN 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1785
      MaxLength       =   15
      TabIndex        =   0
      Top             =   135
      Width           =   1635
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Efectivo ME :"
      Height          =   195
      Index           =   2
      Left            =   825
      TabIndex        =   9
      Top             =   585
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Otro Tipo Cancelación :"
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   8
      Top             =   1350
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Deposito / Cheque :"
      Height          =   195
      Left            =   345
      TabIndex        =   7
      Top             =   945
      Width           =   1440
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Efectivo MN :"
      Height          =   195
      Index           =   0
      Left            =   810
      TabIndex        =   6
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmCuentaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsPago As Recordset

Private Sub cmdOpcion_Click(Index As Integer)
   If Index = 0 Then
   
      
          
        'Inserta Movimiento auditoria
        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTABLA", "TIPO PAGO", "02", sUsuario, "01", "", _
           "TTABLA", "TABLA", "TIPOPAGO", "tCodigo", "Codigo TIPO PAGO", "01", "tValor", "Efectivo MN", txtEfectivoMN.Text, "tIcono", "Efectivo ME", txtEfectivoME.Text)
        If lAuditoria = False Then
            Screen.MousePointer = vbDefault
                Exit Sub
        End If
        'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
               
                 
        'Inserta Movimiento auditoria
        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTABLA", "TIPO PAGO", "02", sUsuario, "03", "", _
           "TTABLA", "TABLA", "TIPOPAGO", "tCodigo", "Codigo TIPO PAGO", "03", "tValor", "Deposito", txtDeposito.Text)
        If lAuditoria = False Then
            Screen.MousePointer = vbDefault
                Exit Sub
        End If
        'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                       
               
           'Inserta Movimiento auditoria
        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTABLA", "TIPO PAGO", "02", sUsuario, "04", "", _
           "TTABLA", "TABLA", "TIPOPAGO", "tCodigo", "Codigo TIPO PAGO", "04", "tValor", "Otro", txtOtro.Text)
        If lAuditoria = False Then
            Screen.MousePointer = vbDefault
                Exit Sub
        End If
        'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                       
               
   
      Cn.Execute "update TTABLA set tValor='" & txtEfectivoMN.Text & "',lreplica=1 where tTabla='TIPOPAGO' and tCodigo='01'"
      Cn.Execute "update TTABLA set tIcono='" & txtEfectivoME.Text & "',lreplica=1 where tTabla='TIPOPAGO' and tCodigo='01'"
      Cn.Execute "update TTABLA set tValor='" & txtDeposito.Text & "',lreplica=1 where tTabla='TIPOPAGO' and tCodigo='03'"
      Cn.Execute "update TTABLA set tValor='" & txtOtro.Text & "',lreplica=1 where tTabla='TIPOPAGO' and tCodigo='04'"
      MsgBox "Registro actualizado", vbExclamation, sMensaje
      Unload Me
   Else
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Centrar Me
   Isql = "select * from vTipoPago"
   Set RsPago = Lib.OpenRecordset(Isql, Cn)
      
   RsPago.Filter = "Codigo='01'"
   txtEfectivoMN.Text = IIf(IsNull(RsPago!CTAMN), "", RsPago!CTAMN)
   txtEfectivoME.Text = IIf(IsNull(RsPago!CTAME), "", RsPago!CTAME)
   
   RsPago.Filter = "Codigo='03'"
   txtDeposito.Text = IIf(IsNull(RsPago!CTAMN), "", RsPago!CTAMN)
   
   RsPago.Filter = "Codigo='04'"
   txtOtro.Text = IIf(IsNull(RsPago!CTAMN), "", RsPago!CTAMN)

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmCuentaPago = Nothing
End Sub


