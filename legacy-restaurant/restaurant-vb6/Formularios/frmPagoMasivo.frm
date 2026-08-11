VERSION 5.00
Begin VB.Form frmPagoMasivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   Icon            =   "frmPagoMasivo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkOtro 
      Height          =   390
      Left            =   150
      TabIndex        =   45
      Top             =   5925
      Width           =   165
   End
   Begin VB.CheckBox chkCheque 
      Height          =   390
      Left            =   150
      TabIndex        =   44
      Top             =   4275
      Width           =   165
   End
   Begin VB.CheckBox chkTarjetas 
      Height          =   315
      Left            =   150
      TabIndex        =   43
      Top             =   1800
      Width           =   165
   End
   Begin VB.CheckBox chkEfectivo 
      Height          =   390
      Left            =   150
      TabIndex        =   42
      Top             =   525
      Width           =   165
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
      Height          =   705
      Index           =   2
      Left            =   2700
      Picture         =   "frmPagoMasivo.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8250
      Width           =   1560
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
      Height          =   705
      Index           =   0
      Left            =   4320
      Picture         =   "frmPagoMasivo.frx":083C
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   8250
      Width           =   1560
   End
   Begin VB.Frame fraOtros 
      Caption         =   "       Otros Pagos "
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
      Height          =   2175
      Left            =   0
      TabIndex        =   25
      Top             =   6000
      Width           =   5925
      Begin VB.CommandButton cmdPago 
         Caption         =   "Recibo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   12
         Left            =   120
         TabIndex        =   28
         Top             =   300
         Width           =   1800
      End
      Begin VB.CommandButton cmdPago 
         Caption         =   "Nota de Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   11
         Left            =   2025
         TabIndex        =   27
         Top             =   300
         Width           =   1800
      End
      Begin VB.CommandButton cmdPago 
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   13
         Left            =   3900
         TabIndex        =   26
         Top             =   300
         Width           =   1800
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Monto :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   750
         TabIndex        =   33
         Top             =   1800
         Width           =   1635
      End
      Begin VB.Label txtDocMonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0EADB&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2805
         TabIndex        =   32
         Top             =   1725
         Width           =   2580
      End
      Begin VB.Label txtDocumento 
         BackColor       =   &H00F0EADB&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   525
         TabIndex        =   31
         Top             =   975
         Width           =   4860
      End
      Begin VB.Label txtReferencia 
         BackColor       =   &H00F0EADB&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2805
         TabIndex        =   30
         Top             =   1350
         Width           =   2580
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         Caption         =   "Referencia :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   750
         TabIndex        =   29
         Top             =   1350
         Width           =   1635
      End
   End
   Begin VB.Frame fraCheque 
      Caption         =   "       Cheque / Depósito "
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
      Height          =   1575
      Left            =   0
      TabIndex        =   20
      Top             =   4350
      Width           =   5925
      Begin VB.CommandButton cmdPago 
         Caption         =   "Número"
         Height          =   555
         Index           =   5
         Left            =   150
         TabIndex        =   22
         Top             =   825
         Width           =   1275
      End
      Begin VB.CommandButton cmdPago 
         Caption         =   "Banco"
         Height          =   555
         Index           =   4
         Left            =   150
         TabIndex        =   21
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label txtNCH 
         BackColor       =   &H00F0EADB&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1500
         TabIndex        =   24
         Top             =   900
         Width           =   4170
      End
      Begin VB.Label txtBCH 
         BackColor       =   &H00F0EADB&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1530
         TabIndex        =   23
         Top             =   390
         Width           =   4170
      End
   End
   Begin VB.Frame fraEfectivo 
      Caption         =   "       Efectivo: "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   0
      TabIndex        =   15
      Top             =   600
      Width           =   5925
      Begin VB.CheckBox chkPagoEfectivoME 
         Caption         =   "Efectivo ME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   3450
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   300
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.CheckBox chkPagoEfectivoMN 
         Caption         =   "Efectivo MN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   1050
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   300
         Width           =   2265
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   7
         Left            =   8940
         Picture         =   "frmPagoMasivo.frx":092E
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   6
         Left            =   8940
         Picture         =   "frmPagoMasivo.frx":11F8
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   795
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   5
         Left            =   8940
         Picture         =   "frmPagoMasivo.frx":1AC2
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1350
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   4
         Left            =   8940
         Picture         =   "frmPagoMasivo.frx":238C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame fraTarjeta 
      Caption         =   "       Tarjetas"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   0
      TabIndex        =   0
      Top             =   1875
      Width           =   5925
      Begin VB.CommandButton cmdPago 
         Caption         =   "Número"
         Height          =   555
         Index           =   7
         Left            =   150
         TabIndex        =   37
         Top             =   1650
         Width           =   825
      End
      Begin VB.CommandButton cmdPago 
         Caption         =   "Fecha Vencimiento"
         Height          =   555
         Index           =   8
         Left            =   3225
         TabIndex        =   36
         Top             =   1650
         Width           =   1275
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   3
         Left            =   8940
         Picture         =   "frmPagoMasivo.frx":2C56
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   2
         Left            =   8940
         Picture         =   "frmPagoMasivo.frx":3520
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1350
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   1
         Left            =   8940
         Picture         =   "frmPagoMasivo.frx":3DEA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   795
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   0
         Left            =   8940
         Picture         =   "frmPagoMasivo.frx":46B4
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   630
         Index           =   8
         Left            =   4425
         TabIndex        =   8
         Top             =   990
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   630
         Index           =   7
         Left            =   3000
         TabIndex        =   7
         Top             =   990
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   630
         Index           =   6
         Left            =   1590
         TabIndex        =   6
         Top             =   990
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   630
         Index           =   5
         Left            =   150
         TabIndex        =   5
         Top             =   990
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   630
         Index           =   4
         Left            =   4425
         TabIndex        =   4
         Top             =   310
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   630
         Index           =   3
         Left            =   3015
         TabIndex        =   3
         Top             =   310
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   630
         Index           =   2
         Left            =   1575
         TabIndex        =   2
         Top             =   310
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         Height          =   630
         Index           =   1
         Left            =   135
         TabIndex        =   1
         Top             =   310
         Width           =   1335
      End
      Begin VB.Label txtVT 
         BackColor       =   &H00F0EADB&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4545
         TabIndex        =   39
         Top             =   1830
         Width           =   1260
      End
      Begin VB.Label txtNT 
         BackColor       =   &H00F0EADB&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1050
         TabIndex        =   38
         Top             =   1800
         Width           =   2085
      End
   End
   Begin VB.Label txtCargo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   2250
      TabIndex        =   14
      Top             =   75
      Width           =   2910
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cargo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   975
      TabIndex        =   13
      Top             =   150
      Width           =   645
   End
End
Attribute VB_Name = "frmPagoMasivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sTarjeta As String
Dim RsTarjeta As Recordset
 
Dim nEN As Double
Dim nEE As Double
Dim nCH As Double
Dim nET As Double
Dim nPT As Double
Dim nDocumento As Double

Dim nAbonoN As Double
  
Dim nSaldo As Double
Dim mTarjeta(8, 3)
Dim sTipoTarjeta As String
Dim sTitulo As String
Dim nIndex As Integer
Dim sCortesia As String
Dim sTipoDocumento As String
Dim sOtroTipoCancelacion As String
Dim sMonDoc As String
Dim nTotalPuntos As Double
Dim sClientePuntos As String
Dim sOtroTipo As String

Dim nRet As Integer
Dim sOperacion As String
Dim sRetorno As String * 512
Dim sClave As String
Dim sMonto As String
Dim xError As String
Dim sRefer As String
Dim nCorrela As String
Dim lEmisor As Boolean
Dim lLoop As Boolean
Dim nContador As Integer
Dim tipoPago As String
Dim tipoTarjeta As String
 Dim tTipoMoneda As String
Dim wPunto As Boolean
Dim sTemp As String
 
Public Sub obtieneSaldo(n As Double)
    nSaldo = n
End Sub

Public Sub pagar()
               'Tipo de Cambio
               Dim nCorrelativo As Integer
                  
               If nTC = 0 Then
                  MsgBox "Tipo de Cambio no ingresado", vbCritical, sMensaje
                  Exit Sub
               End If
               
               'JL Correccion
               If lMCPV Then
                    If MultiCajeroOk = False Then
                        Exit Sub
                    End If
               End If
                             
               Screen.MousePointer = vbHourglass
               wEnter = True
               Dim fFechaPago As Date
               Isql = "Update MDOCUMENTO set " & _
                         "tEstadoDocumento = '02', " & _
                         "nAbono = " & nSaldo & "," & _
                         "fPago = getdate()," & _
                         "nVuelto = 0 " & _
                         ",lreplica=1 where tDocumento ='" & sDocumento & "'"

               Cn.Execute Isql
               nCorrelativo = 1
               Dim nEfectivo As Double
               
               Cn.Execute "delete from DPREPAGO where tDocumento='" & sDocumento & "'"
               
                If xTipo = "Modificacion" Then
                    Cn.Execute "delete from DPAGODOCUMENTO where tDocumento='" & sDocumento & "'"
                    Cn.Execute "UPDATE VC " _
                                & "SET VC.lESTADO_USO = 0 " _
                                & "FROM DPAGODOCUMENTO_VC PVC " _
                                & "INNER JOIN VALE_CONSUMO VC ON PVC.tCod_Vale = VC.tCOD_VALE " _
                                & "WHERE PVC.tDocumento = '" & sDocumento & "' AND " _
                                & "CONVERT(VARCHAR(20),VC.fFECHA_EMISION,23) = CONVERT(VARCHAR(20),GETDATE(),23) "

                    Cn.Execute "DELETE PVC " _
                                & " FROM DPAGODOCUMENTO_VC PVC " _
                                & " INNER JOIN VALE_CONSUMO VC ON PVC.tCod_Vale = VC.tCOD_VALE " _
                                & " WHERE PVC.tDocumento = '" & sDocumento & "' AND " _
                                & " CONVERT(VARCHAR(20),VC.fFECHA_EMISION,23) = CONVERT(VARCHAR(20),GETDATE(),23) "
                End If
               
               Select Case tipoPago
                Case "E"
                            'Efectivo Moneda Nacional
                                 Isql = "insert into DPAGODOCUMENTO " & _
                                      "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, nTipoCambio, nMonto, fRegistro, tUsuario,fDiaContable ) " & _
                                      "Values(   '" & sDocumento & "'," _
                                               & "'" & nCorrelativo & "'," _
                                               & "'" & sTurno & "'," _
                                               & "'01'," _
                                               & "'01'," _
                                               & nTC & ", " _
                                               & nSaldo & ",GETDATE() " _
                                               & ",'" & sUsuario & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "')"
                               Cn.Execute Isql
                 
                Case "T"
                              'Tarjeta
                      Isql = "insert into DPAGODOCUMENTO " & _
                             "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, tReferencia, nTipoCambio, nMonto, npropina, tTarjeta, tNumero, tFechaVencimiento, fRegistro, tUsuario,fDiaContable ) " & _
                             "Values(    '" & sDocumento & "'," _
                                      & "'" & Trim(str(nCorrelativo)) & "'," _
                                      & "'" & sTurno & "'," _
                                      & "'02'," _
                                      & "'01'," _
                                      & "'', " _
                                      & nTC & ", " _
                                      & nSaldo & ", " _
                                      & 0 & ", " _
                                      & "'" & tipoTarjeta & "', " _
                                      & "'', " _
                                      & "'', " _
                                      & "getdate()," _
                                      & "'" & sUsuario & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "')"
                      Cn.Execute Isql
                End Select
               'Liberacion
               
               If xTipo = "" Then
                  Cn.Execute "Update dbo.DPEDIDO set tFacturado ='P', tCortesia='' where tDocumento ='" & sDocumento & "'"
                  Cn.Execute "UPDATE dbo.MPEDIDO set tEstadoPedido='02', lReplica=1 where tCodigoPedido in (select DISTINCT tCodigoPedido FROM DDOCUMENTO where tDocumento='" & sDocumento & "' )  AND TCODIGOPEDIDO NOT IN (SELECT DISTINCT TCODIGOPEDIDO FROM DPEDIDO WHERE TCODIGOPEDIDO IN(select DISTINCT TCODIGOPEDIDO FROM DDOCUMENTO where tDocumento='" & sDocumento & "') AND ISNULL(TFACTURADO,'') <> 'P')"
                  Cn.Execute "UPDATE dbo.MPEDIDO set fLlegada=getdate(), fEntrega = {fn CURDATE()}  where tCodigoPedido in (select DISTINCT tCodigoPedido FROM DDOCUMENTO where tDocumento='" & sDocumento & "' ) and tTipoPedido='02' and isnull(fLlegada ,0)=0"
                  Cn.Execute "Update dbo.TMESA set tEstadoMesa = '04' where tCodigoMesa in (SELECT DISTINCT TMESA FROM MPEDIDO WHERE TCODIGOPEDIDO IN (SELECT DISTINCT TCODIGOPEDIDO FROM DDOCUMENTO WHERE TDOCUMENTO='" & sDocumento & "') AND TCODIGOPEDIDO NOT IN (SELECT DISTINCT TCODIGOPEDIDO FROM DPEDIDO WHERE TCODIGOPEDIDO IN(select DISTINCT TCODIGOPEDIDO FROM DDOCUMENTO where tDocumento='" & sDocumento & "') AND ISNULL(TFACTURADO,'') <> 'P'))"
                  'Juntar Mesa
                  Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='" & sPedido & "')"
               End If
                              
      
               Screen.MousePointer = vbDefault
               wEnter = True
               'Unload Me
               
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkCheque_Click()
       txtBCH.Caption = ""
       txtNCH.Caption = ""
       Me.fraEfectivo.Enabled = False
       Me.fraTarjeta.Enabled = False
       Me.fraOtros.Enabled = False
       
       If chkCheque.value = 1 Then
            tipoPago = "C"
            chkEfectivo.value = 0
            chkTarjetas.value = 0
            chkOtro.value = 0
            Me.fraCheque.Enabled = True
        Else
            tipoPago = ""
            Me.fraCheque.Enabled = False
        
    End If
End Sub

Private Sub chkEfectivo_Click()
        Me.fraTarjeta.Enabled = False
        Me.fraCheque.Enabled = False
        Me.fraOtros.Enabled = False
        Me.chkPagoEfectivoME.value = 0
        Me.chkPagoEfectivoMN.value = 0
        
    If chkEfectivo.value = 1 Then
        chkTarjetas.value = 0
        chkCheque.value = 0
        chkOtro.value = 0
        Me.fraEfectivo.Enabled = True

        tipoPago = "E"
    Else
        tipoPago = ""
        tTipoMoneda = ""
        Me.fraEfectivo.Enabled = False
    End If
End Sub

Private Sub chkOtro_Click()
    txtDocumento.Caption = ""
    txtReferencia.Caption = ""
    txtDocMonto.Caption = ""
    Me.fraEfectivo.Enabled = False
    Me.fraTarjeta.Enabled = False
    Me.fraCheque.Enabled = False

    If chkOtro.value = 1 Then
        tipoPago = "O"
        chkEfectivo.value = 0
        chkCheque.value = 0
        chkTarjetas.value = 0
        Me.fraOtros.Enabled = True
    Else
        tipoPago = ""
        Me.fraOtros.Enabled = False
    End If
End Sub

Private Sub chkPagoEfectivoME_Click()
    If chkPagoEfectivoME.value = 1 Then
        chkPagoEfectivoMN.value = 0
    End If
End Sub

Private Sub chkPagoEfectivoMN_Click()
    If Me.chkPagoEfectivoMN.value = 1 Then
        tipoPago = "E"
        tTipoMoneda = "01"
        chkPagoEfectivoME.value = 0
    End If
End Sub

Private Sub chkTarjetas_Click()
    Dim g As Integer
    For g = 1 To 8
            cmdTarjeta(g).FontBold = False
    Next g
    txtNT.Caption = ""
    txtVT.Caption = ""
    tipoTarjeta = ""
    Me.fraCheque.Enabled = False
    Me.fraOtros.Enabled = False
    Me.fraEfectivo.Enabled = False
    
    If chkTarjetas.value = 1 Then
        tipoPago = "T"
        chkEfectivo.value = 0
        chkCheque.value = 0
        chkOtro.value = 0
        Me.fraTarjeta.Enabled = True
    Else
        tipoPago = ""
        Me.fraTarjeta.Enabled = False
 
    End If
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
        Case 0
                wEnter = False
                Unload Me
        Case 2
                
                tTipoPagoMasivo = tipoPago
                
                If Me.chkEfectivo.value = 1 Then
                    If tTipoMoneda = "" Then
                        MsgBox "Debe seleccionar tipo de Moneda", vbInformation
                        Exit Sub
                    End If
                End If
                
                If Me.chkTarjetas.value = 1 Then
                    If tipoTarjeta = "" Or txtNT.Caption = "" Then
                        MsgBox "Debe indicar los datos necesarios para el pago con Tarjeta", vbInformation
                        Exit Sub
                    End If
                    tTipoTarjetaMasivo = tipoTarjeta
                    tOtroTipoCancelacionMasivo = txtVT.Caption
                    tDocumentoPagoMasivo = txtNT.Caption
                End If
                
                If Me.chkCheque.value = 1 Then
                    If txtNCH.Caption = "" And txtBCH.Caption = "" Then
                        MsgBox "Debe indicar los datos necesarios para el pago con Cheque", vbInformation
                        Exit Sub
                    End If
                    tDocumentoPagoMasivo = txtNCH.Caption
                    tBancoPagoMasivo = txtBCH.Caption
                End If
                
                If Me.chkOtro.value = 1 Then
                    If txtDocumento.Caption = "" Or txtReferencia.Caption = "" Or val(txtDocMonto.Caption) = 0 Then
                        MsgBox "Debe indicar los datos necesarios", vbInformation
                        Exit Sub
                    End If
                
                    If val(txtDocMonto.Caption) <> nSaldo Then
                        MsgBox "El total a pagar debe ser " & txtCargo & " ", vbInformation
                        Exit Sub
                    End If
                    
                    tOtroTipoCancelacionMasivo = sOtroTipoCancelacion
                    If sOtroTipoCancelacion = "000" Then
                    tTipoTarjetaMasivo = Mid(Me.txtDocumento.Caption, 1, 30)
                    Else
                    tTipoTarjetaMasivo = ""
                    End If
                    tDocumentoPagoMasivo = txtReferencia.Caption
                    
                    
                End If
                
                wEnter = True
                Unload Me
    End Select
End Sub

Private Sub cmdPago_Click(Index As Integer)
    Select Case Index
                           
    Case Is = 4  ' Banco
               frmKeyBoard.txtResultado.Text = txtBCH.Caption
               frmKeyBoard.Show vbModal
               If wEnter Then
                  txtBCH.Caption = Mid(sDescrip, 1, 30)
               End If
          
    Case Is = 5  ' Numero
               sTipo = "Numero"
               frmNumPad.Show vbModal
               txtNCH.Caption = IIf(wEnter = True, Mid(sDescrip, 1, 15), txtNCH.Caption)
          
     Case 7
                sTipo = "Numero"
               frmNumPad.Show vbModal
               txtNT.Caption = IIf(wEnter = True, Mid(sDescrip, 1, 16), txtNT.Caption)
          
     Case 8   ' Fecha de Vencimiento
               sTipo = "Fecha"
               frmNumPad.Show vbModal
               
               If wEnter = True Then
                  If Len(sDescrip) = 4 And val(Mid(sDescrip, 1, 2)) <= 12 And val(Mid(sDescrip, 1, 2)) >= 1 Then
                     txtVT.Caption = sDescrip
                  Else
                     MsgBox "Formato no válido (MMAA)", vbInformation, sMensaje
                     txtVT.Caption = ""
                  End If
               Else
                  txtVT.Caption = ""
            End If
    
          Case Is = 11  ' Nota de Crédito
               sTemp = ""
               Isql = "select *, tNotaCredito as Codigo, nVenta as Descripcion from vNotaCredito where tEstadodocumento ='05'"
               Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Nota Credito", 2, "tNotaCredito", 1900, 2, 0, "", _
                                                               "Fecha", 2, "fFecha", 1200, 0, 0, "dd/mm/yyyy", _
                                                               "Cliente", 2, "Cliente", 3900, 0, 0, "", _
                                                               "Monto", 2, "nVenta", 1200, 1, 0, "##,##0.00")
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
               If wEnter Then
                  txtDocumento.Caption = "Nota Credito"
                  sTipoDocumento = sCodigo
                  txtReferencia.Caption = sCodigo
                  sOtroTipoCancelacion = "002"
                  nDocumento = val(sDescrip)
                  sOtroTipo = sTipo
               Else
                  Exit Sub
               End If
               
               txtDocMonto.Caption = Format(nDocumento, "###,###,###,##0.00")
               
              
               
          Case Is = 12  ' Recibos
               sTemp = ""
               Isql = "select *, tRecibo as Codigo, nMonto as Descripcion from vIngreso where tEstadodocumento ='01' and lAnticipo=1 and tmoneda='01'"
               Call ConfGrilla(5, frmBusquedaRapida.grdGrilla, "Recibo", 2, "tRecibo", 1500, 2, 0, "", _
                                                               "Fecha", 2, "fFecha", 1250, 0, 0, "dd/mm/yyyy", _
                                                               "Motivo", 2, "tDescripcion", 3500, 0, 0, "", _
                                                               "Mon", 2, "Moneda", 700, 2, 0, "", _
                                                               "Monto", 2, "nMonto", 1300, 1, 0, "##,##0.00")
               frmBusquedaRapida.nPredeterm = 0
               frmBusquedaRapida.Show vbModal
               
               If wEnter Then
                  txtDocumento.Caption = "Recibo/Anticipo"
                  sTipoDocumento = sCodigo
                  txtReferencia.Caption = sCodigo
                  sOtroTipoCancelacion = "001"
                  sOtroTipo = sTipo
                                    

                  sMonDoc = Calcular("select tMoneda as Codigo from vIngreso where tRecibo='" & sCodigo & "'", Cn)
                  
                  If sMonDoc = "01" Then
                          nDocumento = val(sDescrip)
                  Else
                         nDocumento = val(sDescrip) * nTC
                  End If
                  
               Else
                  Exit Sub
               End If
               txtDocMonto.Caption = Format(nDocumento, "###,###,###,##0.00")
               
                         
          Case Is = 13  ' Varios
               sTemp = ""
               Isql = "select Codigo, Descripcion from vTipoCancelacion Where lActivo='1' and Codigo<>'001' and Codigo<>'002'"
               Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                               "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
                              
               If wEnter = True Then
                  If sCodigo = "000" Then
                     frmKeyBoard.Caption = "Otro Tipo de Cancelación"
                     frmKeyBoard.Show vbModal
                     sOtroTipoCancelacion = "000"
                  Else
                     sOtroTipoCancelacion = sCodigo
                  End If
                  txtDocumento.Caption = sDescrip
               Else
                  Exit Sub
               End If
               
               
               sTipo = ""
               frmNumPad.Show vbModal
               nDocumento = IIf(wEnter = True, sDescrip, nDocumento)
               txtDocMonto.Caption = Format(nDocumento, "###,###,###,##0.00")
            
               
               
               sTipo = "Comanda"
               frmNumPad.Caption = "Referencia"
               frmNumPad.Show vbModal
               sTipoDocumento = IIf(wEnter = True, sDescrip, "")
               txtReferencia.Caption = sTipoDocumento
                           
    End Select
End Sub

Private Sub cmdTarjeta_Click(Index As Integer)
 Dim h As Integer
    For h = 1 To 8
        cmdTarjeta(h).FontBold = False
    Next
   tipoPago = "T"
   tipoTarjeta = ""
   tipoTarjeta = mTarjeta(Index, 1)
   cmdTarjeta(Index).FontBold = True
    
    
End Sub

Private Sub Form_Load()
   wEnter = False
   tipoPago = ""
   tipoTarjeta = ""
   
   frmPagoMasivo.Caption = "Pago de Cuentas por Cobrar"
   sMonDoc = "01"
   Limpia
   
   If lDesactivaNCFP Then
    cmdPago(11).Enabled = False
   Else
    cmdPago(11).Enabled = True
   End If
   
   txtCargo.Caption = Format(nSaldo, "###,###,###,##0.00")

   
   
   'TARJETAS DE CREDITO
   Isql = "select * from TTARJETACREDITO where nBoton>0 and lActivo = 1 Order by nBoton"
   Set RsTarjeta = Lib.OpenRecordset(Isql, Cn)
   Call AsignaTarjeta(8, RsTarjeta, cmdTarjeta())
   ActivaTarjeta False
   'FIN TARJETAS DE CREDITO

   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   
   wPunto = False
   sTemp = "0"
   sDescrip = ""
    Me.chkPagoEfectivoMN.Caption = "Efectivo en " & sMonN
   If sMonE <> "" And sMonN <> sMonE Then
      chkPagoEfectivoME.Caption = "Monto en " & sMonE
   Else
      Me.chkPagoEfectivoME.Visible = False
   End If
   Label3.Caption = "Cargo " & sMonN
End Sub

Public Sub Limpia()
 nEN = 0
 nEE = 0
 nCH = 0
 nET = 0
 nPT = 0
 nDocumento = 0
 nTotalPuntos = 0
 nAbonoN = 0
' nSaldo = nCargo - nAbonoN
End Sub

Public Sub ActivaTarjeta(Activa As Boolean)
   Dim i As Integer
   For i = 1 To 8
       cmdTarjeta(i).Enabled = Not Activa
   Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsTarjeta = Nothing
   Set frmPagoRapido = Nothing
End Sub

Public Sub AsignaTarjeta(nBoton As Integer, RsAsigna As Recordset, cmdBoton As Object)
   Screen.MousePointer = vbHourglass
   Dim i As Integer
   With RsAsigna
        If .RecordCount > 0 Then
           For i = 1 To nBoton
               .MoveFirst
               .Find "nboton = " & Trim(str(i))
               If Not .EOF Then
                  mTarjeta(i, 1) = !tCodigoTarjeta
                  mTarjeta(i, 2) = IIf(IsNull(!tResumido), "", !tResumido)
                  mTarjeta(i, 3) = IIf(IsNull(!lPinPad), 0, !lPinPad)
                  cmdBoton(i).Visible = True
                  cmdBoton(i).Caption = mTarjeta(i, 2)
                Else
                  cmdBoton(i).Visible = False
                End If
           Next i
        Else
           For i = 1 To nBoton
               cmdBoton(i).Visible = False
           Next i
        End If
   End With
   Screen.MousePointer = vbDefault
End Sub

 
'
'
'Private Sub cmdkey_Click(Index As Integer)
'   If sModulo = "ADICION" Then
'      frmMozoUsuario.ReseteaTimer
'   End If
'
'    Select Case Index
'           Case Is = 10 ' Esc
'                wEnter = False
'                Unload Me
'
'           Case Is = 11 ' Supr
'                wPunto = False
'                sTemp = "0"
'                txtResultado.Caption = Format(sTemp, "###,###,###,##0.00")
'                txtTempo.Text = ""
'                txtTempo.SetFocus
'
'                If nPuerto > 0 And sFormulario = "CajaRapida" Then
'                   Visor "Abono :" & sMonN & " " & Right(String(10, " ") & Format(0, "##,##0.00"), 9), "Vuelto:" & sMonN & " " & Right(String(8, " ") & Format(0, "##,##0.00"), 8), nPuerto, "N"
'                End If
'                txtVuelto.Caption = IIf(Val(sTemp) - nSaldo > 0, Format(Val(sTemp) - nSaldo, "###,###,###,##0.00"), "0.00")
'
'           Case Is = 12 'Enter
'                'If Val(sTemp) < nSaldo And Val(sTemp) <> 0 Then
'                If Val(sTemp) < nSaldo Then
'                   Exit Sub
'                End If
'                wEnter = True
'                sDescrip = sTemp
'                tipoPago = "E" ' EFECTIVO
'                pagar
'
'                Unload Me
'
'           Case Is = 13 'Punto
'                If Not wPunto Then
'                   sTemp = sTemp & "."
'                   wPunto = True
'                   txtTempo.SetFocus
'                End If
'
'                If nPuerto > 0 And sFormulario = "CajaRapida" Then
'                   Visor "Abono :" & sMonN & " " & Right(String(10, " ") & Format(Val(sTemp), "##,##0.00"), 9), "Vuelto:" & sMonN & " " & Right(String(8, " ") & Format(IIf(Val(sTemp) - nSaldo > 0, Val(sTemp) - nSaldo, 0), "##,##0.00"), 8), nPuerto, "N"
'                End If
'                txtVuelto.Caption = IIf(Val(sTemp) - nSaldo > 0, Format(Val(sTemp) - nSaldo, "###,###,###,##0.00"), "0.00")
'
'           Case Else
'                If (Not wPunto And Len(Trim(sTemp)) >= 16) Or (wPunto And (Len(Right(Trim(sTemp), Trim(InStr(StrReverse(sTemp), "."))))) > 2 And sTipo = "") Or (wPunto And (Len(Right(Trim(sTemp), Trim(InStr(StrReverse(sTemp), "."))))) > 3 And sTipo = "TC") Then
'                   Beep
'                   txtTempo.SetFocus
'                Else
'                   sTemp = IIf(sTemp = "0", cmdKey(Index).Caption, sTemp & cmdKey(Index).Caption)
'                End If
'                txtResultado.Caption = Format(sTemp, "###,###,###,##0.00")
'                If nPuerto > 0 And sFormulario = "CajaRapida" Then
'                   Visor "Abono :" & sMonN & " " & Right(String(10, " ") & Format(Val(sTemp), "##,##0.00"), 9), "Vuelto:" & sMonN & " " & Right(String(8, " ") & Format(IIf(Val(sTemp) - nSaldo > 0, Val(sTemp) - nSaldo, 0), "##,##0.00"), 8), nPuerto, "N"
'                End If
'                txtVuelto.Caption = IIf(Val(sTemp) - nSaldo > 0, Format(Val(sTemp) - nSaldo, "###,###,###,##0.00"), "0.00")
'                txtVuelto.Caption = IIf(Val(sTemp) - nSaldo > 0, Format(Val(sTemp) - nSaldo, "###,###,###,##0.00"), "0.00")
'                txtTempo.SetFocus
'    End Select
'
'
'End Sub
'
'Private Sub txtTempo_KeyDown(KeyCode As Integer, Shift As Integer)
'   If sModulo = "ADICION" Then
'      frmMozoUsuario.ReseteaTimer
'   End If
'
'   Select Case KeyCode
'   Case 13
'        cmdkey_Click (12)
'   Case 27
'        cmdkey_Click (10)
'   Case 46
'        cmdkey_Click (11)
'   Case 96, 48
'        cmdkey_Click (0)
'   Case 97, 49
'        cmdkey_Click (1)
'   Case 98, 50
'        cmdkey_Click (2)
'   Case 99, 51
'        cmdkey_Click (3)
'   Case 100, 52
'        cmdkey_Click (4)
'   Case 101, 53
'        cmdkey_Click (5)
'   Case 102, 54
'        cmdkey_Click (6)
'   Case 103, 55
'        cmdkey_Click (7)
'   Case 104, 56
'        cmdkey_Click (8)
'   Case 105, 57
'        cmdkey_Click (9)
'   Case 110, 190
'        cmdkey_Click (13)
'   End Select
'End Sub

'diaContable
Public Function obtieneDiaContable() As Date
   Dim oComando As New clsComando
   Dim DiaContable As Date
   Dim rst1 As New ADODB.Recordset
   Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("usp_GenObtieneDiaContable", Cn) Then
                     Set oComando = Nothing
                     Exit Function
                  End If
                  
                  oComando.CreateParameter "@lDiaContable", adBoolean, adParamInput, 1, lDiaContable
                  oComando.CreateParameter "@sHoraCierre", adVarChar, adParamInput, 5, tHoraCierreDiaContable
                  oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, sUsuario
                 oComando.CreateParameter "@fDiaContable", adDBDate, adParamOutput, 10, DiaContable
                If Not oComando.GetParamOK Then
                   Set oComando = Nothing
                   Exit Function
                End If
                    Set rst1 = oComando.GetSP()
                obtieneDiaContable = oComando.GetParameterValue("@fDiaContable")
End Function
'diaContable




