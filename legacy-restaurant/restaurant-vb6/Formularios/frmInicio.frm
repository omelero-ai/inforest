VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmInicio 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inicio de Turno"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frmInicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Caption         =   "Tipo de Cambio Oficial"
      Height          =   555
      Index           =   5
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Frame fraMontos 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Montos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   6900
      Begin VB.TextBox txtAbonoE 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   2520
         TabIndex        =   13
         Top             =   1440
         Width           =   1365
      End
      Begin VB.TextBox txtAbonoN 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   2475
         TabIndex        =   12
         Top             =   765
         Width           =   1365
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
         Caption         =   "Abono MN"
         Height          =   555
         Index           =   3
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   555
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
         Caption         =   "Abono ME"
         Height          =   555
         Index           =   4
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Anterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0073B43C&
         Height          =   195
         Left            =   4005
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Inicial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0073B43C&
         Height          =   195
         Left            =   5505
         TabIndex        =   21
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Abono Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0073B43C&
         Height          =   195
         Left            =   2580
         TabIndex        =   20
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label txtAnteriorN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   3930
         TabIndex        =   19
         Top             =   765
         Width           =   1365
      End
      Begin VB.Label txtAnteriorE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   3930
         TabIndex        =   18
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label txtSaldoN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5385
         TabIndex        =   17
         Top             =   765
         Width           =   1365
      End
      Begin VB.Label txtSaldoE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5400
         TabIndex        =   16
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label txtME 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1605
         TabIndex        =   15
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label txtMN 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1605
         TabIndex        =   14
         Top             =   765
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Caption         =   "Tipo de Cambio"
      Height          =   555
      Index           =   2
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   1395
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
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
      Height          =   555
      Index           =   1
      Left            =   4380
      Picture         =   "frmInicio.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Caption         =   "Apertura"
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
      Index           =   0
      Left            =   5730
      Picture         =   "frmInicio.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   1275
   End
   Begin MSComCtl2.Animation aniVideo 
      Height          =   540
      Left            =   120
      TabIndex        =   23
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   953
      _Version        =   393216
      FullWidth       =   49
      FullHeight      =   36
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Caption         =   "Tipo de Cambio"
      Height          =   555
      Index           =   6
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1800
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label txtTC2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   120
      TabIndex        =   28
      Top             =   1320
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label txtTC3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   120
      TabIndex        =   27
      Top             =   1920
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label txtTCO 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1965
   End
   Begin VB.Label lblProceso 
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizando datos. Este proceso puede tomar algunos minutos."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   585
      Left            =   900
      TabIndex        =   24
      Top             =   3780
      Visible         =   0   'False
      Width           =   3225
   End
   Begin VB.Label txtTC 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1965
   End
   Begin VB.Label TxtFecha 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   1395
   End
   Begin VB.Label txtUsuario 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1965
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0073B43C&
      Height          =   195
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0073B43C&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   660
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTc As Recordset
Dim RsTurno As Recordset
Dim wAgrega As Boolean
Dim nAnteriorN As Double
Dim nAnteriorE As Double
Dim nAbonoN As Double
Dim nAbonoE As Double
Dim nMontoSN As Double
Dim nMontoSE As Double
 

Private Sub cmdOpcion_Click(Index As Integer)
   Dim nCorrela As String
   Dim xMensaje As String
   Dim wPasa As Boolean
   
   Select Case Index
          Case Is = 0 ' Aperturar
                If pais = "003" Then
                     If (sMonE <> "" Or sMonE2 <> "" Or sMonE3 <> "") And (sMonN <> sMonE Or sMonN <> sMonE2 Or sMonN <> sMonE3) And (nTC = 0 Or nTC2 = 0 Or nTC3 = 0) Then
                        MsgBox "Tipo de cambio no ingresado", vbExclamation, sMensaje
                        Exit Sub
                     End If
                Else
                     If sMonE <> "" And sMonN <> sMonE And nTC = 0 Then
                        MsgBox "Tipo de cambio no ingresado", vbExclamation, sMensaje
                        Exit Sub
                     End If
                End If
               '-----------------
               If RsTurno.RecordCount = 0 Then
                  If MsgBox("Seguro de Aperturar el Turno?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                     Exit Sub
                  End If
               Else
                  If RsTurno!lCierre = True Then
                     If MsgBox("Seguro de Aperturar el Turno?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                        Exit Sub
                     Else
                        wInicioActivaApertura = True
                     End If
                  Else
                     If MsgBox("Seguro de Re Aperturar el Turno?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                        Exit Sub
                     Else
                        wInicioActivaApertura = True
                     End If
                  End If
               End If
               
               
               If wAgrega Then
                  Dim oComando As clsComando
                  Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("spIns_TipoCambio", Cn) Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  '---CESAR - tipocambio
                  oComando.CreateParameter "@nTc", adDouble, adParamInput, 0, nTC
                  oComando.CreateParameter "@tUSUARIO", adVarChar, adParamInput, 15, sUsuario
                  oComando.CreateParameter "@nTco", adDouble, adParamInput, 0, nTCO
                  oComando.CreateParameter "@nTc2", adDouble, adParamInput, 0, nTC2
                  oComando.CreateParameter "@nTc3", adDouble, adParamInput, 0, nTC3
                  If Not oComando.GetParamOK Then
                  '------------
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  If Not oComando.ExecSP Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
               End If
                              
               If RsTurno.RecordCount = 0 Then
                  wPasa = True
               Else
                  wPasa = RsTurno!lCierre
               End If
                  
               If wPasa Then
                  'Obtiene el Correlativo
                  nCorrela = Calcular("select max(tTurno) as Codigo from MTURNO where substring(tTurno,1,2)= substring(ltrim(str(year(getdate()))),3,2)", Cn)
                  If IsNull(nCorrela) Or Mid(nCorrela, 1, 2) <> Mid(Trim(str(Year(FechaServidor()))), 3, 2) Then
                     nCorrela = Mid(Trim(str(Year(FechaServidor()))), 3, 2) & "00000001"
                  Else
                     nCorrela = Mid(Trim(str(Year(FechaServidor()))), 3, 2) & Lib.Correlativo(Mid(nCorrela, 3, 8), 8)
                  End If
                  
                  
                  Isql = "insert into MTURNO( " & _
                         "tTurno, tCaja, tSalon, fInicial, tUsuario, lCierre, nMontoIN, nMontoIE) " & _
                         "values ('" & nCorrela & "', " & _
                                 "'" & sCaja & "', " & _
                                 "'" & sSalon & "', " & _
                                 "getdate() , " & _
                                 "'" & sUsuario & "', " & _
                                        0 & ", " & _
                                 nAbonoN & ", " & _
                                 nAbonoE & ")"
                  Cn.Execute Isql
                  sTurno = nCorrela
                  If lAlmacenRemoto = True Then
                     actualizaDatosSistemaAlmacen
                  End If
                  
               Else
                  sTurno = RsTurno!tTurno
                   
                  Isql = "update MTURNO set " & _
                         "tUsuario ='" & sUsuario & "', " & _
                         "nMontoIN = " & nAbonoN & ", " & _
                         "nMontoIE = " & nAbonoE & " " & _
                         "where tTurno ='" & sTurno & "'"
                   Cn.Execute Isql
               End If
               
               
              'ENVIA LOS RESUMENES DE BOL Y FAC - GCAA 23012023
                
                
                If (lFEBiz) And (lFEBizOffLine) Then
                     If INSERTA_FE_INFOREST("", 3, DateTime.Now) = False Then
                                                 
                     End If
                    
                     If INSERTA_FE_INFOREST("", 4, DateTime.Now) = False Then
                                                 
                     End If
                End If
                
                If pais = "000" Then
                    If lFacturacionE = True And lFEBiz = False And lFEOfisis = False And lFEpape = False And lFECarbajal = False And lFESpring = False And lFEGesa = False Then
                        INSERTAFE Format(FechaServidor() - 1, "yyyymmdd"), "", 3, ""
                        INSERTAFE Format(FechaServidor() - 2, "yyyymmdd"), "", 3, ""
                        INSERTAFE Format(FechaServidor() - 3, "yyyymmdd"), "", 3, ""
                        INSERTAFE Format(FechaServidor() - 4, "yyyymmdd"), "", 3, ""
                    End If
                End If
               
               ActivaInicio (True)
               wInicio = True
               Unload Me
          
          Case Is = 1 ' Cancelar
               Unload Me

          Case Is = 2 ' Tipo de Cambio
              sTipo = "TC"
               frmNumPad.Show vbModal
               If wEnter Then
                  txtTC.Caption = Format(sDescrip, "###,###,##0.000")
                  nTC = val(sDescrip)
                If pais = "003" Then
                    txtTCO.Caption = Format(sDescrip, "###,###,##0.000")
                    nTCO = val(sDescrip)
                End If
               End If
          Case Is = 3 ' Abono MN
               sTipo = ""
               frmNumPad.Show vbModal
               nAbonoN = IIf(wEnter = True, sDescrip, nAbonoN)
               txtAbonoN.Text = Format(nAbonoN, "###,###,###,##0.00")
               txtSaldoN.Caption = Format(nAbonoN + nAnteriorN, "###,###,###,##0.00")
          
          Case Is = 4 ' Abono ME
               sTipo = ""
               frmNumPad.Show vbModal
               nAbonoE = IIf(wEnter = True, sDescrip, nAbonoE)
               txtAbonoE.Text = Format(nAbonoE, "###,###,###,##0.00")
               txtSaldoE.Caption = Format(nAbonoE + nAnteriorE, "###,###,###,##0.00")
         '---CESAR tipo cambio Oficial
         Case Is = 5
                sTipo = "TC"
               frmNumPad.Show vbModal
               If wEnter Then
                  If pais = "003" Then
                   txtTC2.Caption = Format(sDescrip, "###,###,##0.000")
                   nTC2 = val(sDescrip)
                  Else
                    txtTCO.Caption = Format(sDescrip, "###,###,##0.000")
                    nTCO = val(sDescrip)
                  End If
               End If
         Case Is = 6
               sTipo = "TC"
               frmNumPad.Show vbModal
               If wEnter Then
                  If pais = "003" Then
                   txtTC3.Caption = Format(sDescrip, "###,###,##0.000")
                   nTC3 = val(sDescrip)
                  End If
               End If
          '---------------
   End Select
End Sub

Private Sub Form_Activate()
   If CDate(Format(txtFecha.Caption, "short date")) > FechaServidor() Then
      MsgBox "Error : Ha querido ingresar un turno con fecha anterior", vbCritical, sMensaje
      Unload Me
   End If
End Sub

Private Sub Form_Load()
  
   Centrar Me
   On Error Resume Next
   aniVideo.Open App.path & "\bmps\FileMove.avi"
'    aniVideo.Open App.Path & "\bmps\peya.avi"
'    aniVideo.Visible = True
'    aniVideo.AutoPlay = True

'INFORESTAPP
   Dim TurnoCompartido As Integer
   
   If Calcular("SELECT COUNT(*) AS Codigo FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'TCAJA' AND COLUMN_NAME = 'lTurnoCompartido'", Cn) > 0 Then
      TurnoCompartido = 1
   Else
      TurnoCompartido = 0
   End If
   
   If lMCPV Then
      Isql = "select * from MTURNO where tUsuario ='" & sUsuario & "' order by tTurno"

   ElseIf TurnoCompartido = 1 Then
      Isql = "select * from MTURNO where tCaja ='" & sCaja & "' and tUsuario ='" & sUsuario & "' order by tTurno"
   Else
      Isql = "select * from MTURNO where tCaja ='" & sCaja & "' order by tTurno"
   End If
   Set RsTurno = Lib.OpenRecordset(Isql, Cn)
   
   If RsTurno.RecordCount = 0 Then
      nAbonoN = 0
      nAbonoE = 0
      nAnteriorN = 0
      nAnteriorE = 0
      nMontoSN = 0
      nMontoSE = 0
      
      txtAbonoN.Text = Format(nAbonoN, "###,###,##0.00")
      txtAbonoE.Text = Format(nAbonoE, "###,###,##0.00")
      txtAnteriorN.Caption = Format(nAnteriorN, "###,###,##0.00")
      txtAnteriorE.Caption = Format(nAnteriorE, "###,###,##0.00")
      txtSaldoN.Caption = Format(nMontoSN, "###,###,##0.00")
      txtSaldoE.Caption = Format(nMontoSE, "###,###,##0.00")
            
      txtFecha.Caption = FechaServidor()
      Me.Caption = "Apertura de Turno"
   Else
      RsTurno.MoveLast
                  
      If Not RsTurno!lCierre = True Then
         txtFecha.Caption = IIf(IsNull(RsTurno!finicial), Now, RsTurno!finicial)
         Me.Caption = "Re Apertura de Turno"
         nAbonoN = IIf(IsNull(RsTurno!nMontoIN), 0, RsTurno!nMontoIN)
         nAbonoE = IIf(IsNull(RsTurno!nMontoIE), 0, RsTurno!nMontoIE)
         nAnteriorN = 0
         nAnteriorE = 0
         nMontoSN = nAbonoN + nAnteriorN
         nMontoSE = nAbonoE + nAnteriorE
                           
         txtAbonoN.Text = Format(nAbonoN, "###,###,##0.00")
         txtAbonoE.Text = Format(nAbonoE, "###,###,##0.00")
         txtAnteriorN.Caption = Format(nAnteriorN, "###,###,##0.00")
         txtAnteriorE.Caption = Format(nAnteriorE, "###,###,##0.00")
         txtSaldoN.Caption = Format(nMontoSN, "###,###,##0.00")
         txtSaldoE.Caption = Format(nMontoSE, "###,###,##0.00")
      Else
         Me.Caption = "Apertura de Turno"
         txtFecha.Caption = FechaServidor()
         
         nAbonoN = 0
         nAbonoE = 0
         nAnteriorN = 0
         nAnteriorE = 0
         
         nMontoSN = nAbonoN + nAnteriorN
         nMontoSE = nAbonoE + nAnteriorE
         
         txtAbonoN.Text = Format(nAbonoN, "###,###,##0.00")
         txtAbonoE.Text = Format(nAbonoE, "###,###,##0.00")
         txtAnteriorN.Caption = Format(nAnteriorN, "###,###,##0.00")
         txtAnteriorE.Caption = Format(nAnteriorE, "###,###,##0.00")
         txtSaldoN.Caption = Format(nMontoSN, "###,###,##0.00")
         txtSaldoE.Caption = Format(nMontoSE, "###,###,##0.00")
      End If
   End If
   
   If sMonE <> "" And sMonN <> sMonE And pais <> "003" Then  'If sMonE <> "" And sMonN <> sMonE  Then
      Set RsTc = Lib.OpenRecordset("SELECT * From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
           
      If RsTc.EOF Then
         nTC = 0
         nTC2 = 0
         nTC3 = 0
         nTCO = 0
         wAgrega = True
      Else
         nTC = IIf(IsNull(RsTc!nVenta), 0, IIf(IsNull(RsTc!nVenta), 0, RsTc!nVenta))
         nTC2 = IIf(IsNull(RsTc!nVenta2), 0, IIf(IsNull(RsTc!nVenta2), 0, RsTc!nVenta2))
         nTC3 = IIf(IsNull(RsTc!nVenta3), 0, IIf(IsNull(RsTc!nVenta3), 0, RsTc!nVenta3))
         'CESAR-----para mostrar el tipo cambio si ya esta registrado
         nTCO = IIf(IsNull(RsTc!nOficial), 0, IIf(IsNull(RsTc!nOficial), 0, RsTc!nOficial))
         wAgrega = False
         
         If nTC = 0 Then: wAgrega = True
         If nTCO = 0 Then: wAgrega = True
         
      End If
      If nTC = 0 And Not lInfhotel Then
         cmdOpcion(2).Visible = True
      Else
         cmdOpcion(2).Visible = False
      End If
      
      If nTCO = 0 And Not lInfhotel Then
         '---CESAR
         cmdOpcion(5).Visible = True
      Else
         '---CESAR
         cmdOpcion(5).Visible = False
      End If
      
      txtTC.Caption = Format(nTC, "#,###,##0.000")
      txtTC2.Caption = Format(nTC2, "#,###,##0.000")
      txtTC3.Caption = Format(nTC3, "#,###,##0.000")
      
      '---CESAR tipo cambio
      txtTCO.Caption = Format(nTCO, "#,###,##0.000")
      txtMN.Caption = sMonN
      txtME.Caption = sMonE
   ElseIf (sMonE <> "" Or sMonE2 <> "" Or sMonE3 <> "") And (sMonN <> sMonE Or sMonN <> sMonE2 Or sMonN <> sMonE3) And pais = "003" Then   'If sMonE <> "" And sMonN <> sMonE  Then
      Set RsTc = Lib.OpenRecordset("SELECT * From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
           
      If RsTc.EOF Then
         nTC = 0
         nTC2 = 0
         nTC3 = 0
         nTCO = 0
         wAgrega = True
      Else
         nTC = IIf(IsNull(RsTc!nVenta), 0, IIf(IsNull(RsTc!nVenta), 0, RsTc!nVenta))
         nTC2 = IIf(IsNull(RsTc!nVenta2), 0, IIf(IsNull(RsTc!nVenta2), 0, RsTc!nVenta2))
         nTC3 = IIf(IsNull(RsTc!nVenta3), 0, IIf(IsNull(RsTc!nVenta3), 0, RsTc!nVenta3))
         'CESAR-----para mostrar el tipo cambio si ya esta registrado
         nTCO = IIf(IsNull(RsTc!nOficial), 0, IIf(IsNull(RsTc!nOficial), 0, RsTc!nOficial))
         wAgrega = False
         
         If nTC = 0 Then: wAgrega = True
         If nTC3 = 0 Then: wAgrega = True
         If nTC2 = 0 Then: wAgrega = True
         If nTCO = 0 Then: wAgrega = True
         
      End If
      If nTC = 0 And Not lInfhotel Then
         cmdOpcion(2).Visible = True
      Else
         cmdOpcion(2).Visible = False
      End If

      If nTCO = 0 And Not lInfhotel Then
         '---CESAR
         cmdOpcion(5).Visible = True
      Else
         '---CESAR
         cmdOpcion(5).Visible = False
      End If
      
      If pais = "003" Then
        If nTC2 = 0 And Not lInfhotel Then
           cmdOpcion(5).Visible = True
        Else
           cmdOpcion(5).Visible = False
        End If
        If nTC3 = 0 And Not lInfhotel Then
           cmdOpcion(6).Visible = True
        Else
           cmdOpcion(6).Visible = False
        End If
      End If
      txtTC.Caption = Format(nTC, "#,###,##0.000")
      txtTC2.Caption = Format(nTC2, "#,###,##0.000")
      txtTC3.Caption = Format(nTC3, "#,###,##0.000")
      
      '---CESAR tipo cambio
      txtTCO.Caption = Format(nTCO, "#,###,##0.000")
      txtMN.Caption = sMonN
      txtME.Caption = sMonE
   Else
      cmdOpcion(2).Visible = False
      cmdOpcion(4).Visible = False
      '---CESAR tipo cambio
      cmdOpcion(5).Visible = False
      cmdOpcion(6).Visible = False
      txtMN.Caption = sMonN
      txtME.Visible = False
      txtAbonoE.Visible = False
      txtAbonoE.Visible = False
      txtSaldoE.Visible = False
      txtAnteriorE.Visible = False
      txtTC.Caption = "0.000"
      txtTC3.Caption = "0.000"
      txtTC2.Caption = "0.000"
      nTC = 1
      nTC2 = 1
      nTC3 = 1
      '---CESAR tipo cambio
      txtTCO.Caption = "0.000"
      nTCO = 1
      
   End If
   txtUsuario.Caption = sUsuario
   
    'TIPO CAMBIO
    If pais = "002" Then
        'Label2.Visible = False
        'Label8.Visible = False
        'Label9.Visible = False
        txtTC.Visible = False
        txtTCO.Visible = False
        cmdOpcion(2).Visible = False
        cmdOpcion(5).Visible = False
    ElseIf pais = "003" Then ' Argentina
        'Label8.Visible = False
        'Label7.Visible = True
        'Label9.Visible = False
        'Label2.Caption = "T. Cambio " & sMonE & ":"
        'Label8.Caption = "T. Cambio " & sMonE2 & ":"
        'Label7.Caption = "T. Cambio " & sMonE3 & ":"
        fraMontos.Height = 1260
        fraMontos.Top = 2520
'        cmdOpcion(2).Visible = True
'        cmdOpcion(5).Visible = True
'        cmdOpcion(6).Visible = True
        txtTC3.Visible = True
        txtTC2.Visible = True
        cmdOpcion(2).Caption = "Tipo Cambio " & Calcular("select isnull(tresumido,'') as codigo from vmoneda where codigo='02'", Cn)
        cmdOpcion(5).Caption = "Tipo Cambio " & Calcular("select isnull(tresumido,'') as codigo from vmoneda where codigo='03'", Cn)
        cmdOpcion(6).Caption = "Tipo Cambio " & Calcular("select isnull(tresumido,'') as codigo from vmoneda where codigo='04'", Cn)
    ElseIf pais = "005" Then 'España
            txtTCO.Visible = False
            txtTC.Visible = False
            txtTC2.Visible = False
            
            cmdOpcion(4).Visible = False
            txtME.Visible = False
            txtAbonoE.Visible = False
            txtAnteriorE.Visible = False
            txtSaldoE.Visible = False
    Else
        'Label2.Visible = True
        'Label8.Visible = True
        'Label9.Visible = True
        txtTC.Visible = True
        txtTCO.Visible = True
       
        If VisibleTC = 1 Then
            cmdOpcion(2).Visible = True
            cmdOpcion(5).Visible = True
            cmdOpcion(2).Caption = "Tipo cambio " & Calcular("select isnull(tresumido,'') as codigo from vmoneda where codigo='02'", Cn)
            cmdOpcion(5).Caption = "Tipo cambio oficial " & Calcular("select isnull(tresumido,'') as codigo from vmoneda where codigo='02'", Cn)
        End If
    End If
    
    
'    If pais = "003" Then ' Argentina
'        Label8.Visible = False
'        Label9.Visible = False
'
'        fraMontos.Height = 1260
'        fraMontos.Top = 2520
'        cmdOpcion(6).Visible = True
'        txtTC3.Visible = True
'        txtTC2.Visible = True
'        cmdOpcion(2).Caption = "Tipo Cambio " & Calcular("select isnull(tresumido,'') as codigo from vmoneda where codigo='02'", Cn)
'        cmdOpcion(5).Caption = "Tipo Cambio " & Calcular("select isnull(tresumido,'') as codigo from vmoneda where codigo='03'", Cn)
'        cmdOpcion(6).Caption = "Tipo Cambio " & Calcular("select isnull(tresumido,'') as codigo from vmoneda where codigo='04'", Cn)
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsTc = Nothing
   Set RsTurno = Nothing
   Set frmInicio = Nothing
End Sub

Public Sub actualizaDatosSistemaAlmacen()
    Screen.MousePointer = vbHourglass
    aniVideo.Visible = True
    aniVideo.AutoPlay = True
    lblProceso.Visible = True
    If VerificaConexionAlmacenRemoto = True Then
        CnAlmacenRemoto.Execute "sp_ActualizaReceta"
        CargaTablasAlmacenRemoto
    End If
    aniVideo.AutoPlay = False
    aniVideo.Visible = False
    lblProceso.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Function VerificaConexionAlmacenRemoto() As Boolean
    On Error GoTo Err:
    If sRutaAlmacenRemoto <> "" And sMDBAlmacenRemoto <> "" Then
        Set CnAlmacenRemoto = New ADODB.Connection
        CnAlmacenRemoto.Provider = "SQLOLEDB"
        CnAlmacenRemoto.CursorLocation = adUseServer
        CnAlmacenRemoto.ConnectionString = "User ID=" & sUserName & _
                ";password=" & sUserPassword & _
                ";Data Source=" & sRutaAlmacenRemoto & _
                ";Initial Catalog=" & sMDBAlmacenRemoto
        CnAlmacenRemoto.CommandTimeout = 0
        CnAlmacenRemoto.Open
            If CnAlmacenRemoto.State Then
                VerificaConexionAlmacenRemoto = True
            Else
                VerificaConexionAlmacenRemoto = False
            End If
    End If
    Exit Function
Err:
    VerificaConexionAlmacenRemoto = False
End Function

 
