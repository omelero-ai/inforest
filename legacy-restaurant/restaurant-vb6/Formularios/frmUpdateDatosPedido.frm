VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmActualizarPedidos 
   Caption         =   "Actualizar Cortesia / Canal de Venta"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Canal de Venta"
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   8895
      Begin VB.CommandButton BtnCliente 
         Caption         =   "..."
         Height          =   375
         Left            =   1330
         TabIndex        =   17
         Top             =   1250
         Visible         =   0   'False
         Width           =   500
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Modificar"
         Height          =   855
         Left            =   6840
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo CbCanal 
         Height          =   315
         Left            =   1440
         TabIndex        =   11
         Top             =   840
         Width           =   5115
         _ExtentX        =   9022
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
      Begin VB.Label txtCliente 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   16
         Top             =   1200
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Label TxtCodCliente 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   15
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Cliente Delivery:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label txtodCanal 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label txtCanal 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   8
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label4 
         Caption         =   "Canal de Venta Actual:"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Modificar por:"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Motivo Cortesia"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8895
      Begin VB.CommandButton BtnActualizar1 
         Caption         =   "Modificar"
         Height          =   855
         Left            =   6840
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo CbCortesia 
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   840
         Width           =   5115
         _ExtentX        =   9022
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
      Begin VB.Label txtCodCortesia 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label txtCortesia 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   6
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label2 
         Caption         =   "Modificar por:"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cortesia Actual:"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmActualizarPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsResult As Recordset
Dim RsCortesia As Recordset
Dim RsCanal As Recordset

Private Sub BtnActualizar_Click()


End Sub

Private Sub BtnActualizar1_Click()
'aqui 20012022 GCAA
   
               
If Supervisor("31") = False Then
   Exit Sub

               
Else
    If txtCodCortesia.Caption = "" Then
        MsgBox "El documnento no tiene asignado cortesia, por lo cual no se le puede asignar una cortesia", vbCritical, sMensaje
        Exit Sub
    End If

   Dim oComandox As clsComando
   Set oComandox = New clsComando

   Set oComandox = New clsComando
    If Not oComandox.CreateCmdSp("usp_ActualizarCabPeDoc", Cn) Then
       Set oComandox = Nothing
       Exit Sub
    End If
    
    oComandox.CreateParameter "@opcion", adVarChar, adParamInput, 20, "1"
    oComandox.CreateParameter "@param1", adVarChar, adParamInput, 50, sDocumento
    oComandox.CreateParameter "@param2", adVarChar, adParamInput, 50, CbCortesia.BoundText
    oComandox.CreateParameter "@param3", adVarChar, adParamInput, 50, txtCodCortesia.Caption
    oComandox.CreateParameter "@param4", adVarChar, adParamInput, 50, sUsuario
    oComandox.CreateParameter "@param5", adVarChar, adParamInput, 50, ""
    oComandox.CreateParameter "@param6", adVarChar, adParamInput, 50, ""
    
    If Not oComandox.GetParamOK Then
       Set oComandox = Nothing
       Exit Sub
    End If

    If Not oComandox.ExecSP Then
       Set oComandox = Nothing
       Exit Sub
     End If
   'fin 06082021 GCAA

 MsgBox "Registro Modificado Correctamente", vbInformation, sMensaje
iniciar

End If

End Sub

Private Sub BtnCliente_Click()
  '  frmBusquedaDelivery.txtResultado.Caption = txtTelefono.Caption
 frmBusquedaDelivery.Show vbModal
    
    txtCliente.Caption = txtCliente.Caption & " " & sDescrip
    TxtCodCliente.Caption = sCodigo
    
End Sub

Private Sub CbCanal_Click(Area As Integer)
If CbCanal.BoundText = "02" Then
    BtnCliente.Visible = True
    TxtCodCliente.Visible = True
    txtCliente.Visible = True
Else
     BtnCliente.Visible = False
    TxtCodCliente.Visible = False
    txtCliente.Visible = False
End If
End Sub

Private Sub Command1_Click()
'aqui 20012022 GCAA
 If Supervisor("32") = False Then
 Exit Sub
   
   Else

   Dim oComandox As clsComando
   Set oComandox = New clsComando

   Set oComandox = New clsComando
    If Not oComandox.CreateCmdSp("usp_ActualizarCabPeDoc", Cn) Then
       Set oComandox = Nothing
       Exit Sub
    End If
    
    oComandox.CreateParameter "@opcion", adVarChar, adParamInput, 20, "2"
    oComandox.CreateParameter "@param1", adVarChar, adParamInput, 50, sDocumento
    oComandox.CreateParameter "@param2", adVarChar, adParamInput, 50, CbCanal.BoundText
    oComandox.CreateParameter "@param3", adVarChar, adParamInput, 50, TxtCodCliente.Caption
    oComandox.CreateParameter "@param4", adVarChar, adParamInput, 50, txtodCanal.Caption
    oComandox.CreateParameter "@param5", adVarChar, adParamInput, 50, sUsuario
    oComandox.CreateParameter "@param6", adVarChar, adParamInput, 50, ""
    
    If Not oComandox.GetParamOK Then
       Set oComandox = Nothing
       Exit Sub
    End If

    If Not oComandox.ExecSP Then
       Set oComandox = Nothing
       Exit Sub
     End If
   'fin 06082021 GCAA
   BtnCliente.Visible = False
    TxtCodCliente.Visible = False
    TxtCodCliente.Visible = False
   MsgBox "Registro Modificado Correctamente", vbInformation, sMensaje
iniciar
End If
End Sub

Private Sub Form_Load()




iniciar
End Sub




Sub iniciar()


Isql = "select distinct   t2.tCodigoPedido ,t1.tCortesia,isnull(t5.Descripcion,'') DescripcionCortesia , t3.tTipoPedido, t4.Descripcion from mDOCUMENTO t1 left join DDOCUMENTO t2  on t1.tDocumento = t2.tDocumento left join MPEDIDO t3   on t2.tCodigoPedido = t3.tCodigoPedido left join vTipoPedido t4     on t3.tTipoPedido = t4.Codigo  left join vCortesia t5   on t5.Codigo = t1.tCortesia where t1.tDocumento ='" & frmDocumentoCorrelativo.grdGrilla.Columns(2).value & "'"

Set RsResult = Lib.OpenRecordset(Isql, Cn)

txtCodCortesia.Caption = IIf(IsNull(RsResult!tCortesia), "", RsResult!tCortesia)
txtCortesia.Caption = IIf(IsNull(RsResult!DescripcionCortesia), "", RsResult!DescripcionCortesia)

txtodCanal.Caption = IIf(IsNull(RsResult!tTipoPedido), "", RsResult!tTipoPedido)
txtCanal.Caption = IIf(IsNull(RsResult!Descripcion), "", RsResult!Descripcion)


  With CbCortesia
         Isql = "select * from vCortesia where lActivo = 1 order by Codigo"
         Set RsCortesia = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsCortesia
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With

  With CbCanal
         Isql = "select * from vTipoPedido where lActivo = 1 order by Codigo"
         Set RsCanal = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsCanal
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
End Sub
