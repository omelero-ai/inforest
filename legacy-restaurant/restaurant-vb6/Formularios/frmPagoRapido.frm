VERSION 5.00
Begin VB.Form frmPagoRapido 
   BackColor       =   &H00E2DEDB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos"
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "frmPagoRapido.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdKey 
      BackColor       =   &H00F2FAED&
      Caption         =   "Otros Medios de Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   14
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7560
      Width           =   2325
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F0F0F0&
      Caption         =   "Otros: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007EC642&
      Height          =   1500
      Left            =   60
      TabIndex        =   38
      Top             =   7080
      Width           =   3000
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   7
         Left            =   8940
         Picture         =   "frmPagoRapido.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   6
         Left            =   8940
         Picture         =   "frmPagoRapido.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1350
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   5
         Left            =   8940
         Picture         =   "frmPagoRapido.frx":149E
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   795
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   4
         Left            =   8940
         Picture         =   "frmPagoRapido.frx":1D68
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
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
         Height          =   630
         Left            =   3120
         Picture         =   "frmPagoRapido.frx":2632
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.PictureBox cmdKeyPicture 
      BackColor       =   &H00E2DEDB&
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   12
      Left            =   3240
      Picture         =   "frmPagoRapido.frx":2724
      ScaleHeight     =   735
      ScaleWidth      =   2775
      TabIndex        =   37
      Top             =   7920
      Width           =   2775
   End
   Begin VB.PictureBox cmdOpcion 
      BackColor       =   &H00E2DEDB&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3240
      Picture         =   "frmPagoRapido.frx":4189
      ScaleHeight     =   735
      ScaleWidth      =   2775
      TabIndex        =   36
      Top             =   7080
      Width           =   2775
   End
   Begin VB.Frame fraEfectivo 
      BackColor       =   &H00F0F0F0&
      Caption         =   "Pago en Efectivo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007EC642&
      Height          =   5175
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   6060
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   1170
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4185
         Width           =   1725
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   1185
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3330
         Width           =   855
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3330
         Width           =   855
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   3
         Left            =   2895
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3330
         Width           =   855
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   4
         Left            =   1185
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2475
         Width           =   855
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   5
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2475
         Width           =   855
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   6
         Left            =   2895
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2475
         Width           =   855
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   7
         Left            =   1185
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1620
         Width           =   855
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   8
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1620
         Width           =   855
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   9
         Left            =   2895
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1620
         Width           =   855
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "Esc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   10
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1620
         Width           =   990
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "Sup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2560
         Index           =   11
         Left            =   3750
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2475
         Width           =   990
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H007EC642&
         Caption         =   "Enter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1715
         Index           =   12
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3000
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdKey 
         BackColor       =   &H00F2FAED&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   13
         Left            =   2895
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4185
         Width           =   855
      End
      Begin VB.TextBox txtTempo 
         Height          =   375
         Left            =   3915
         TabIndex        =   0
         Top             =   3765
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vuelto"
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
         Left            =   4440
         TabIndex        =   34
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Abono"
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
         Left            =   4440
         TabIndex        =   33
         Top             =   795
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   4455
         TabIndex        =   32
         Top             =   405
         Width           =   645
      End
      Begin VB.Label txtCargo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
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
         Height          =   405
         Left            =   1215
         TabIndex        =   31
         Top             =   315
         Width           =   3150
      End
      Begin VB.Label txtVuelto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
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
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   1215
         TabIndex        =   30
         Top             =   1125
         Width           =   3150
      End
      Begin VB.Label txtResultado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0F0&
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
         ForeColor       =   &H00808000&
         Height          =   405
         Left            =   1215
         TabIndex        =   29
         Top             =   720
         Width           =   3150
      End
   End
   Begin VB.Frame fraGrilla 
      BackColor       =   &H00F0F0F0&
      Caption         =   "Pago con: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007EC642&
      Height          =   1815
      Left            =   45
      TabIndex        =   2
      Top             =   5160
      Width           =   6060
      Begin VB.CommandButton cmdOpcion_OLD 
         BackColor       =   &H00FFFFFF&
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
         Height          =   630
         Left            =   3120
         Picture         =   "frmPagoRapido.frx":574A
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2520
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         BackColor       =   &H00FFFFFF&
         Height          =   630
         Index           =   1
         Left            =   375
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         BackColor       =   &H00FFFFFF&
         Height          =   630
         Index           =   2
         Left            =   1695
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         BackColor       =   &H00FFFFFF&
         Height          =   630
         Index           =   3
         Left            =   3015
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         BackColor       =   &H00FFFFFF&
         Height          =   630
         Index           =   4
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         BackColor       =   &H00FFFFFF&
         Height          =   630
         Index           =   5
         Left            =   375
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1020
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         BackColor       =   &H00FFFFFF&
         Height          =   630
         Index           =   6
         Left            =   1695
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1020
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         BackColor       =   &H00FFFFFF&
         Height          =   630
         Index           =   7
         Left            =   3015
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1020
         Width           =   1335
      End
      Begin VB.CommandButton cmdTarjeta 
         BackColor       =   &H00FFFFFF&
         Height          =   630
         Index           =   8
         Left            =   4335
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1020
         Width           =   1335
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   0
         Left            =   8940
         Picture         =   "frmPagoRapido.frx":583C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   1
         Left            =   8940
         Picture         =   "frmPagoRapido.frx":6106
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   795
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   2
         Left            =   8940
         Picture         =   "frmPagoRapido.frx":69D0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1350
         Width           =   1215
      End
      Begin VB.CommandButton cmdNavegar 
         Height          =   555
         Index           =   3
         Left            =   8940
         Picture         =   "frmPagoRapido.frx":729A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmPagoRapido"
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

Dim tMotivoCancelacion As String
Dim tReferencia As String
 
Dim wPunto As Boolean
Dim sTemp As String
Public MontoBilleteRapido As Double


Private Sub cmdKeyPicture_Click(Index As Integer)
    cmdkey_Click 12
End Sub

Private Sub cmdOpcion_Click()
    wEnter = False
    Unload Me
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
                         "nAbono = " & txtResultado.Caption & "," & _
                         "fPago = getdate()," & _
                         "nVuelto = " & txtVuelto.Caption & " " & _
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
                                      "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, nTipoCambio, nMonto, fRegistro, tUsuario,fDiaContable,ndolar ) " & _
                                      "Values(   '" & sDocumento & "'," _
                                               & "'" & nCorrelativo & "'," _
                                               & "'" & sTurno & "'," _
                                               & "'01'," _
                                               & "'01'," _
                                               & nTC & ", " _
                                               & nSaldo & ",GETDATE() " _
                                               & ",'" & sUsuario & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "', " & sTemp & ")"
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
                      
                Case "VA"
                        
                      'Varios
                              
                      Isql = "insert into DPAGODOCUMENTO " & _
                             "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, tReferencia, nTipoCambio, nMonto, npropina, tTarjeta, tBanco, tNumero, tFechaVencimiento, fRegistro, tUsuario,fDiaContable ) " & _
                             "Values(    '" & sDocumento & "'," _
                                      & "'" & Trim(str(nCorrelativo)) & "'," _
                                      & "'" & sTurno & "'," _
                                      & "'04'," _
                                      & "'01'," _
                                      & "'', " _
                                      & nTC & ", " _
                                      & nSaldo & ", " _
                                      & 0 & ", " _
                                      & "'', " _
                                      & "'" & tMotivoCancelacion & "', " _
                                      & "'" & tReferencia & "', " _
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
                              
               If CD Then
                    Call ModifcarEstadoDeliveryCabecera(sDocumento)
               End If
               Screen.MousePointer = vbDefault
               wEnter = True
               'Unload Me
               
End Sub

Private Sub cmdTarjeta_Click(Index As Integer)

    If val(sTemp) < nSaldo Then
       Exit Sub
    End If

    tipoPago = "T"
    tipoTarjeta = ""
    tipoTarjeta = mTarjeta(Index, 1)
                
    If nPuerto > 0 And sFormulario = "CajaRapida" Then
       Dim sss As String
       sss = mTarjeta(Index, 2)
       Visor "Pago con Tarjeta", sss, nPuerto, "N"
    End If

    pagar
    Unload Me
End Sub

Private Sub Form_Load()
   wEnter = False
   tipoPago = ""
   tipoTarjeta = ""
   
   If lCancelacion Then
      cmdOpcion.Enabled = False
   End If
   
   frmPagoRapido.Caption = "Cancelación del Documento " & Format(sDocumento, "@-@@@@@-@@@@@@@@@")
   sMonDoc = "01"
   Limpia
   
    If xTipo = "Modificacion" Then
        txtCargo.Caption = Format(nSaldo, "###,###,###,##0.00")
    Else
        'Obtiene el total del documento para la cancelación!!!!
'        If lModuloPago = "CajaRapida" Then
'            nSaldo = nTotalPR
'        Else
'            nSaldo = nTotalPR
'        End If
        nSaldo = Calcular("select nventa as codigo from MDOCUMENTO where tdocumento= '" & sDocumento & "'", Cn)
        txtCargo.Caption = Format(nSaldo, "###,###,###,##0.00")
    End If
   
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
   txtResultado.Caption = Format(sTemp, "###,###,###,##0.00")
   If MontoBilleteRapido > 0 Then
        cmdkey_Click 999
   End If
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
 nSaldo = nCargo - nAbonoN
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

Private Sub ModifcarEstadoDeliveryCabecera(ByVal qDocumento As String) 'pp
    On Error GoTo ErrorHandler
    'Central Delivery-Motorizado--------------------------------------------------pp
    Isql = "Select Distinct P.tCodigoPedidoCD from DDocumento as D Inner Join MPedido AS P On D.tCodigoPedido = P.tCodigoPedido Where tDocumento = '" + qDocumento + "'"
    Dim RsCodigoPCD As ADODB.Recordset
    Set RsCodigoPCD = Lib.OpenRecordset(Isql, Cn)
    If Not RsCodigoPCD.EOF Then
        If Not IsNull(RsCodigoPCD!tCodigoPedidoCD) Then
'            Call ModifcarEstadoDeliveryCabecera(RsCodigoPCD!tCodigoPedidoCD, "3", txtMotorizado.Caption)
             
            Dim CnCD As Connection
            'Configuración
            Set CnCD = New Connection
            CnCD.Provider = "SQLOLEDB"
            CnCD.CursorLocation = adUseServer
            CnCD.ConnectionString = "User ID=" & sUserName & _
            ";password=" & sUserPassword & _
            ";Data Source=" & sRutaCD & _
            ";Initial Catalog=" & sMDBCD
            CnCD.CommandTimeout = 250
            CnCD.Open
                
            CnCD.Execute "usp_CD_Modificar_EstadoDelivery_Cabecera 4, '" + RsCodigoPCD!tCodigoPedidoCD + "',''"
            CnCD.Close
            
        End If
    End If
    ''''''''''''''''''''''''''''''
    Exit Sub
ErrorHandler:
    MsgBox (Err.Description)
End Sub

Private Sub cmdBorra_AfterClick()
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   sTemp = "0"
   wPunto = False
   txtResultado.Caption = Format(sTemp, "###,###,###,##0.00")
End Sub

Private Sub cmdEnter_AfterClick()
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   wEnter = True
   Unload Me
End Sub

Private Sub cmdEsc_AfterClick()
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   wEnter = False
   Unload Me
End Sub

Private Sub cmdkey_Click(Index As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

    Select Case Index
           Case Is = 10 ' Esc
                wEnter = False
                Unload Me
                
           Case Is = 11 ' Supr
                wPunto = False
                sTemp = "0"
                txtResultado.Caption = Format(sTemp, "###,###,###,##0.00")
                txtTempo.Text = ""
                txtTempo.SetFocus
                
                If nPuerto > 0 And sFormulario = "CajaRapida" Then
                   Visor "Abono :" & sMonN & " " & Right(String(10, " ") & Format(0, "##,##0.00"), 9), "Vuelto:" & sMonN & " " & Right(String(8, " ") & Format(0, "##,##0.00"), 8), nPuerto, "N"
                End If
                txtVuelto.Caption = IIf(val(sTemp) - nSaldo > 0, Format(val(sTemp) - nSaldo, "###,###,###,##0.00"), "0.00")
                                               
           Case Is = 12 'Enter
                'If Val(sTemp) < nSaldo And Val(sTemp) <> 0 Then
                If val(sTemp) < nSaldo Then
                   Exit Sub
                End If
                wEnter = True
                sDescrip = sTemp
                tipoPago = "E" ' EFECTIVO
                pagar
                
                Unload Me
           
           Case Is = 13 'Punto
                If Not wPunto Then
                   sTemp = sTemp & "."
                   wPunto = True
                   txtTempo.SetFocus
                End If
                
                If nPuerto > 0 And sFormulario = "CajaRapida" Then
                   Visor "Abono :" & sMonN & " " & Right(String(10, " ") & Format(val(sTemp), "##,##0.00"), 9), "Vuelto:" & sMonN & " " & Right(String(8, " ") & Format(IIf(val(sTemp) - nSaldo > 0, val(sTemp) - nSaldo, 0), "##,##0.00"), 8), nPuerto, "N"
                End If
                txtVuelto.Caption = IIf(val(sTemp) - nSaldo > 0, Format(val(sTemp) - nSaldo, "###,###,###,##0.00"), "0.00")
                
                
           Case Is = 14 'Otras formas de pago
           
                Dim tUsuActua As String
                
                If lPassOtrosPagos Then
                    If Supervisor("28") = False Then
                        MsgBox "Clave no permitida", vbExclamation, sMensaje
                        Exit Sub
                    End If
                End If
                

               tMotivoCancelacion = ""
               tReferencia = ""
               
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
                  tMotivoCancelacion = sDescrip
               Else
                  Exit Sub
               End If
                              
               tipoPago = "VA"
               tipoTarjeta = ""
               
               
               'Cliente frecuente - linea de credito
               'Dim lActivaCredito As String
               'lActivaCredito = Calcular("select isnull(nvalor,'0') Codigo from vTipoCancelacion Where lActivo='1' and Codigo='" & sOtroTipoCancelacion & "'", Cn)
               
               frmNumPad.Caption = "Referencia"
               frmNumPad.Show vbModal
               sTipoDocumento = IIf(wEnter = True, sDescrip, "")
               tReferencia = sTipoDocumento
               
               pagar
               Unload Me
               
                
           Case Is = 999 'pago billete rapido
                sTemp = MontoBilleteRapido
                txtResultado.Caption = Format(sTemp, "###,###,###,##0.00")
                txtVuelto.Caption = IIf(val(sTemp) - nSaldo > 0, Format(val(sTemp) - nSaldo, "###,###,###,##0.00"), "0.00")
                txtVuelto.Caption = IIf(val(sTemp) - nSaldo > 0, Format(val(sTemp) - nSaldo, "###,###,###,##0.00"), "0.00")
                'txtTempo.SetFocus
           Case Else
                If (Not wPunto And Len(Trim(sTemp)) >= 16) Or (wPunto And (Len(Right(Trim(sTemp), Trim(InStr(StrReverse(sTemp), "."))))) > 2 And sTipo = "") Or (wPunto And (Len(Right(Trim(sTemp), Trim(InStr(StrReverse(sTemp), "."))))) > 3 And sTipo = "TC") Then
                   Beep
                   txtTempo.SetFocus
                Else
                   sTemp = IIf(sTemp = "0", cmdKey(Index).Caption, sTemp & cmdKey(Index).Caption)
                End If
                txtResultado.Caption = Format(sTemp, "###,###,###,##0.00")
                If nPuerto > 0 And sFormulario = "CajaRapida" Then
                   Visor "Abono :" & sMonN & " " & Right(String(10, " ") & Format(val(sTemp), "##,##0.00"), 9), "Vuelto:" & sMonN & " " & Right(String(8, " ") & Format(IIf(val(sTemp) - nSaldo > 0, val(sTemp) - nSaldo, 0), "##,##0.00"), 8), nPuerto, "N"
                End If
                txtVuelto.Caption = IIf(val(sTemp) - nSaldo > 0, Format(val(sTemp) - nSaldo, "###,###,###,##0.00"), "0.00")
                txtVuelto.Caption = IIf(val(sTemp) - nSaldo > 0, Format(val(sTemp) - nSaldo, "###,###,###,##0.00"), "0.00")
                txtTempo.SetFocus
    End Select
    

End Sub




Private Sub txtTempo_KeyDown(KeyCode As Integer, Shift As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   Select Case KeyCode
   Case 13
        cmdkey_Click (12)
   Case 27
        cmdkey_Click (10)
   Case 46
        cmdkey_Click (11)
   Case 96, 48
        cmdkey_Click (0)
   Case 97, 49
        cmdkey_Click (1)
   Case 98, 50
        cmdkey_Click (2)
   Case 99, 51
        cmdkey_Click (3)
   Case 100, 52
        cmdkey_Click (4)
   Case 101, 53
        cmdkey_Click (5)
   Case 102, 54
        cmdkey_Click (6)
   Case 103, 55
        cmdkey_Click (7)
   Case 104, 56
        cmdkey_Click (8)
   Case 105, 57
        cmdkey_Click (9)
   Case 110, 190
        cmdkey_Click (13)
   End Select
End Sub

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


