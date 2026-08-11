VERSION 5.00
Begin VB.Form frmInicioAdicion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apertura de Turno"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "frmInicioAdicion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMontos 
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
      Height          =   1500
      Left            =   90
      TabIndex        =   9
      Top             =   1080
      Width           =   6900
      Begin VB.CommandButton cmdCajon 
         Caption         =   "Cajón de dinero"
         Height          =   555
         Left            =   3960
         TabIndex        =   14
         Top             =   630
         Width           =   1275
      End
      Begin VB.CommandButton cmdAbono 
         Caption         =   "Abono MN"
         Height          =   555
         Left            =   210
         TabIndex        =   10
         Top             =   555
         Width           =   1275
      End
      Begin VB.Label txtCajon 
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
         Top             =   810
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cajón de Dinero"
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
         Left            =   4095
         TabIndex        =   15
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   2520
         TabIndex        =   13
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label txtAbono 
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
         Left            =   2385
         TabIndex        =   12
         Top             =   765
         Width           =   1365
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
         TabIndex        =   11
         Top             =   765
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdTC 
      Caption         =   "Tipo de Cambio"
      Height          =   555
      Left            =   2850
      TabIndex        =   8
      Top             =   495
      Width           =   1275
   End
   Begin VB.CommandButton cmdSalir 
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
      Index           =   0
      Left            =   4380
      Picture         =   "frmInicioAdicion.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2625
      Width           =   1275
   End
   Begin VB.CommandButton cmdApertura 
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
      Left            =   5730
      Picture         =   "frmInicioAdicion.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2625
      Width           =   1275
   End
   Begin VB.Label txtTC 
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
      Height          =   285
      Left            =   1395
      TabIndex        =   5
      Top             =   765
      Width           =   1365
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
      Left            =   1395
      TabIndex        =   4
      Top             =   405
      Width           =   1365
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
      Left            =   1395
      TabIndex        =   3
      Top             =   45
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha :"
      Height          =   195
      Left            =   750
      TabIndex        =   2
      Top             =   450
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Cambio :"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   810
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Usuario :"
      Height          =   195
      Left            =   660
      TabIndex        =   0
      Top             =   90
      Width           =   630
   End
End
Attribute VB_Name = "frmInicioAdicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTc As Recordset
Dim wAgrega As Boolean
Dim nAbono As Double
Public wTurno As Double
Public wMozo As Double
Dim sCajon As String

Private Sub cmdAbono_Click()
    sTipo = ""
    frmNumPad.Show vbModal
    nAbono = IIf(wEnter = True, sDescrip, nAbono)
    txtAbono.Caption = Format(nAbono, "###,###,###,##0.00")
    wEnter = False
End Sub

Private Sub cmdApertura_Click()
    Dim nCorrela As String
    If sMonE <> "" And sMonN <> sMonE And nTC = 0 Then
       MsgBox "Tipo de cambio no ingresado", vbExclamation, sMensaje
       Exit Sub
    End If

    If MsgBox("Seguro de Aperturar el Turno?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
       Exit Sub
    End If
           
    If wAgrega Then
       Dim oComando As clsComando
       Set oComando = New clsComando
       If Not oComando.CreateCmdSp("spIns_TipoCambio", Cn) Then
          Set oComando = Nothing
          Exit Sub
       End If
       oComando.CreateParameter "@nTc", adDouble, adParamInput, 0, nTC
       oComando.CreateParameter "@tUSUARIO", adVarChar, adParamInput, 15, Mid(sUsuario, 1, 15)
       If Not oComando.GetParamOK Then
          Set oComando = Nothing
          Exit Sub
       End If
       If Not oComando.ExecSP Then
          Set oComando = Nothing
          Exit Sub
       End If
    End If
    
    If wTurno Then
       'Obtiene el Correlativo
       nCorrela = Calcular("select max(tTurno) as Codigo from MTURNO where substring(tTurno,1,2)= substring(ltrim(str(year(getdate()))),3,2)", Cn)
       If IsNull(nCorrela) Or Mid(nCorrela, 1, 2) <> Mid(Trim(Str(Year(Date))), 3, 2) Then
          nCorrela = Mid(Trim(Str(Year(Date))), 3, 2) & "00000001"
       Else
          nCorrela = Mid(Trim(Str(Year(Date))), 3, 2) & Lib.Correlativo(Mid(nCorrela, 3, 8), 8)
       End If
       
       Isql = "insert into MTURNO( " & _
              "tTurno, tCaja, tSalon, fInicial, tUsuario, lCierre, nMontoIN, nMontoIE) " & _
              "values ('" & nCorrela & "', " & _
                      "'" & sCaja & "', " & _
                      "'" & sSalon & "', " & _
                      "getdate() , " & _
                      "'" & Mid(sUsuario, 1, 15) & "', " & _
                             0 & ", " & _
                      0 & ", " & _
                      0 & ")"
       Cn.Execute Isql
       sTurno = nCorrela
    End If
    
    If wMozo Then
       nCorrela = Calcular("select max(tItem) as Codigo from MTURNOCAJERO where tTurno='" & sTurno & "'", Cn)
       If IsNull(nCorrela) Then
          nCorrela = "001"
       Else
          nCorrela = Lib.Correlativo(nCorrela, 3)
       End If
    
       Isql = "insert into MTURNOCAJERO( " & _
              "tTurno, tItem, tMozo, fFechaInicio, lCierre, nAbono, tCajon) " & _
              "values ('" & sTurno & "', " & _
                      "'" & nCorrela & "', " & _
                      "'" & sMozo & "', " & _
                      "getdate() , 0, " & _
                      nAbono & ", '" & _
                      sCajon & "')"
       Cn.Execute Isql
       Cn.Execute "update TTABLA set nBoton=1 where tTabla='CAJON' and tCodigo='" & sCajon & "'"
    End If
    wEnter = True
    Unload Me

End Sub

Private Sub cmdCajon_Click()
    sTemp = ""
    Isql = "select * from vCajon where lActivo = 1 and tCaja='" & sCaja & "' and isnull(SW,0)=0 Order by Descripcion"
    Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                    "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
    frmBusquedaRapida.nPredeterm = 1
    frmBusquedaRapida.Show vbModal
    If wEnter = True Then
       txtCajon.Caption = sDescrip
       sCajon = sCodigo
    End If
End Sub

Private Sub cmdSalir_Click(Index As Integer)
   wEnter = False
   Unload Me
End Sub

Private Sub cmdTC_Click()
    sTipo = "TC"
    frmNumPad.Show vbModal
    If wEnter Then
       txtTC.Caption = Format(sDescrip, "###,###,##0.000")
       nTC = Val(sDescrip)
       wEnter = False
    End If
End Sub

Private Sub Form_Activate()
   If CDate(Format(txtFecha.Caption, "short date")) > Date Then
      MsgBox "Error : Ha querido ingresar un turno con fecha anterior", vbCritical, sMensaje
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   wEnter = False
   sCajon = ""
   Centrar Me
   If wTurno Then
      nAbono = 0
   Else
      nAbono = Calcular("select nMontoIN as Codigo from MTURNO where tTurno='" & sTurno & "'", Cn)
   End If
   'Isql = "select * from MTURNO where tCaja ='" & sCaja & "' order by tTurno"
   'Set RsTurno = Lib.OpenRecordset(Isql, Cn)
      
   txtAbono.Caption = Format(nAbono, "###,###,##0.00")
   txtFecha.Caption = Date
   txtUsuario.Caption = sVar1
   
   If sMonE <> "" And sMonN <> sMonE Then
      Set RsTc = Lib.OpenRecordset("SELECT * From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
      
      If RsTc.EOF Then
         nTC = 0
         wAgrega = True
      Else
         nTC = IIf(IsNull(RsTc!nVenta), 0, IIf(IsNull(RsTc!nVenta), 0, RsTc!nVenta))
         wAgrega = False
      End If
      If nTC = 0 And Not lInfhotel Then
         cmdTC.Visible = True
      Else
         cmdTC.Visible = False
      End If
      txtTC.Caption = Format(nTC, "#,###,##0.000")
      txtMN.Caption = sMonN
   Else
      cmdTC.Visible = False
      txtMN.Caption = sMonN
      txtTC.Caption = "0.000"
      nTC = 1
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsTc = Nothing
   Set frmInicioAdicion = Nothing
End Sub
