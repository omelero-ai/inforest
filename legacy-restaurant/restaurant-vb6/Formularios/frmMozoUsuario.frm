VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMozoUsuario 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Acceso "
   ClientHeight    =   9000
   ClientLeft      =   6585
   ClientTop       =   3705
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C000&
   Icon            =   "frmMozoUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMozoUsuario.frx":000C
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmocualtamensaje 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   300
      Left            =   0
      TabIndex        =   21
      Top             =   7080
      Width           =   5400
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   8640
      TabIndex        =   19
      Top             =   11760
      Width           =   6255
   End
   Begin VB.PictureBox imgLogoPic 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2420
      Left            =   7440
      ScaleHeight     =   2385
      ScaleWidth      =   2505
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   2530
   End
   Begin VB.CommandButton cmdConsultaSaldo 
      BackColor       =   &H00FFFFFF&
      DragIcon        =   "frmMozoUsuario.frx":18832
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11160
      Picture         =   "frmMozoUsuario.frx":18C74
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   6360
      MaxLength       =   25
      PasswordChar    =   "*"
      TabIndex        =   14
      Top             =   3600
      Width           =   3960
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   19320
      Top             =   9960
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Aceptar"
      Height          =   800
      Index           =   0
      Left            =   2640
      Picture         =   "frmMozoUsuario.frx":193B6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   11760
      Width           =   1170
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PassWord"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   2
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   12000
      Width           =   1170
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   3
      Left            =   720
      Picture         =   "frmMozoUsuario.frx":194B8
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   11640
      Width           =   840
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4770
      Left            =   1800
      TabIndex        =   2
      Top             =   12720
      Width           =   11760
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Mensajes "
      Height          =   1845
      Left            =   12960
      TabIndex        =   5
      Top             =   12240
      Width           =   7770
   End
   Begin VB.Timer Timer_girar 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   16830
      Top             =   9765
   End
   Begin VB.Timer Timer_LlenaRecordSet 
      Interval        =   1000
      Left            =   16350
      Top             =   9765
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0C0&
      DragIcon        =   "frmMozoUsuario.frx":1A1FA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16680
      Picture         =   "frmMozoUsuario.frx":1A63C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   10800
      Width           =   765
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00C0C0C0&
      Height          =   750
      Index           =   1
      Left            =   20520
      MaskColor       =   &H000000C0&
      Picture         =   "frmMozoUsuario.frx":1AA7E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10440
      Width           =   765
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   15870
      Top             =   9765
   End
   Begin VB.Timer timSalida 
      Interval        =   3000
      Left            =   15390
      Top             =   9765
   End
   Begin MCI.MMControl mmControl 
      Height          =   375
      Left            =   17280
      TabIndex        =   0
      Top             =   9780
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   661
      _Version        =   393216
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PlayVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   18360
      TabIndex        =   3
      Top             =   11760
      Width           =   7605
   End
   Begin VB.Frame frmcontenido 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   9495
      Left            =   15480
      TabIndex        =   10
      Top             =   120
      Width           =   6735
      Begin VB.Frame Frame4 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   735
         Left            =   360
         TabIndex        =   11
         Top             =   3840
         Visible         =   0   'False
         Width           =   4695
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Actualización Disponible"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   13
            Top             =   0
            Width           =   4695
         End
         Begin VB.Label LblVer 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   12
            Top             =   360
            Width           =   4695
         End
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1935
      Left            =   0
      TabIndex        =   20
      Top             =   7080
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   3413
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lblconexion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0073B43C&
      Height          =   255
      Left            =   6360
      TabIndex        =   22
      Top             =   8760
      Width           =   4575
   End
   Begin VB.Image Imagemensaje 
      Height          =   2460
      Left            =   -120
      Picture         =   "frmMozoUsuario.frx":1AB70
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   5520
   End
   Begin VB.Image ImagePais 
      Height          =   465
      Left            =   3650
      Stretch         =   -1  'True
      Top             =   340
      Width           =   645
   End
   Begin VB.Image Image1 
      Height          =   1140
      Left            =   0
      Picture         =   "frmMozoUsuario.frx":1EB1D
      Top             =   0
      Width           =   4500
   End
   Begin VB.Image imgingresar 
      Height          =   735
      Left            =   6360
      Top             =   6240
      Width           =   4695
   End
   Begin VB.Image ImgLogo 
      Height          =   2415
      Left            =   7440
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label txtFecha 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0073B43C&
      Height          =   285
      Left            =   7680
      TabIndex        =   17
      Top             =   420
      Width           =   705
   End
   Begin VB.Label lblHora 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0073B43C&
      Height          =   285
      Left            =   6240
      TabIndex        =   16
      Top             =   420
      Width           =   705
   End
   Begin VB.Image imgOpcion 
      Height          =   1095
      Index           =   3
      Left            =   8160
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Image imgNewProceso 
      Height          =   735
      Index           =   0
      Left            =   6360
      Top             =   7080
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   8985
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5385
   End
   Begin VB.Image ImageHelp 
      Height          =   735
      Left            =   10920
      Top             =   7920
      Width           =   735
   End
   Begin VB.Image imgOpcion 
      Height          =   735
      Index           =   1
      Left            =   11280
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imgNewProceso 
      Height          =   300
      Index           =   2
      Left            =   17880
      Picture         =   "frmMozoUsuario.frx":23332
      Stretch         =   -1  'True
      Top             =   10800
      Width           =   2205
   End
   Begin VB.Image imgOpcion 
      Height          =   615
      Index           =   2
      Left            =   10440
      Top             =   3480
      Width           =   600
   End
   Begin VB.Image ImgLogo2 
      Height          =   1305
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   11640
      Width           =   3000
   End
   Begin VB.Image Image2 
      Height          =   1005
      Left            =   5040
      Picture         =   "frmMozoUsuario.frx":32D40
      Stretch         =   -1  'True
      Top             =   11760
      Width           =   3030
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "www.infhotel.com.pe"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BF8801&
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   12480
      Width           =   3030
   End
End
Attribute VB_Name = "frmMozoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsMozo As Recordset
Dim rsMensajeCocina As Recordset
Dim rsparam As Recordset

Dim fso As Object
Dim sql_mensaje As String
Dim nroElementos As Integer
   Dim RsTc As New ADODB.Recordset
Dim i As Integer
Dim marcacion As Integer
Private Const PI = 3.14159265
Dim inicio As Boolean

Private Sub ValidacionEntrarConLicencia()
    If lHARDkey Then
        '----------Verifica Llave HK----------------------------------
        Dim verif As Boolean
        verif = hk.VerificaConexion
                        
        If verif = False Then
            Dim str As String
            str = hk.IniciaConexion(InfhotelHK.Adicion)
            If str <> "" Then
                MsgBox str, vbCritical, "Aviso"
                End
            End If
        End If
        '--------------------------------------------------------------
    End If
End Sub

Private Sub cmdConsultaSaldo_Click()
    frmConsultaSaldo.Show vbModal
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
           Case Is = 0 ' Aceptar
                'TIPO CAMBIO
                  If lMCPV = False Then
                    If pais <> "002" Then
                        Dim RsTc As New ADODB.Recordset
                        Set RsTc = Lib.OpenRecordset("SELECT * From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
                
                        If RsTc.EOF Then
                         nTC = 0
                        Else
                         nTC = IIf(IsNull(RsTc!nVenta), 0, IIf(IsNull(RsTc!nVenta), 0, RsTc!nVenta))
                        End If
                
                        wInicio = False
                        If nTC = 0 Then
                         MsgBox "Error: No se ha ingresado el Tipo de Cambio", vbCritical, sMensaje
                         txtPassword.Text = ""
                         txtPassword.SetFocus
                         Exit Sub
                        End If
                    End If
                 End If
                
                If lHARDkey Then
                    ValidacionEntrarConLicencia
                End If
                If txtPassword.Text = "" Then
                   MsgBox "Ingrese su password", vbExclamation, sMensaje
                   txtPassword.SetFocus
                   Exit Sub
                Else
                    RsMozo.MoveFirst
                    Do While Not RsMozo.EOF
                       If Desencapsula(IIf(IsNull(RsMozo!tvalor), "", RsMozo!tvalor)) = UCase(txtPassword.Text) Or (IIf(IsNull(RsMozo!tBandaMagnetica), "", RsMozo!tBandaMagnetica) = Encapsula(UCase(Extrae(txtPassword.Text))) And Encapsula(UCase(Extrae(txtPassword.Text))) <> "") Then
                             sVar1 = RsMozo!tResumido
                             If RsMozo!nTamano = 1 Then
                                Dim sCambio As String
                                MsgBox "Por motivo de seguridad, deberá cambiar su password", vbCritical, sMensaje
                                
                                frmPassword.Caption = "Ingrese su nueva clave"
                                frmPassword.Show vbModal
                                If Not wEnter Then
                                   Exit Sub
                                End If
                                sCambio = sDescrip
                                frmPassword.Caption = "Confirme su nueva clave"
                                frmPassword.Show vbModal
                                If Not wEnter Then
                                   Exit Sub
                                End If
                                If sCambio <> sDescrip Or sDescrip = "" Then
                                   MsgBox "Confirmación erronea, no se realizó el cambio", vbCritical, sMensaje
                                   Exit Sub
                                End If
                                RsMozo!nTamano = 0
                                RsMozo!tvalor = Encapsula(sDescrip)
                             End If
                          sPassword = UCase(txtPassword.Text)
                          sMozo = RsMozo!codigo
                          lSomelier = IIf(IsNull(RsMozo!nValor), 0, RsMozo!nValor)
                          txtPassword.Text = ""
                          If lMCPV Then
                            lImportacionReservas = False
                             If sModulo = "INFOREST" Then
                                sUsuario = sVar1
                                Unload Me
                             Else
                             
                                lImportacionReservas = IIf(IsNull(RsMozo!lImportacionReservas), False, RsMozo!lImportacionReservas)
                                'audirotia
                                
                                registroAccesoAuditoria "I", sVar1
                                
                                'auditoria


                                Me.Timer_LlenaRecordSet.Enabled = False
                                Me.Timer_LlenaRecordSet.Interval = 0
                                frmCargoMozo.Show vbModal
                                Me.Timer_LlenaRecordSet.Enabled = True
                                Me.Timer_LlenaRecordSet.Interval = 1000
                                
                                
                             End If
                          Else
                          
                                lImportacionReservas = IIf(IsNull(RsMozo!lImportacionReservas), False, RsMozo!lImportacionReservas)
                                
                                'audirotia
                                registroAccesoAuditoria "I", sVar1
                                
                                'auditoria
                                Me.Timer_LlenaRecordSet.Enabled = False
                                Me.Timer_LlenaRecordSet.Interval = 0
                                frmCargoMozo.Show vbModal
                                Me.Timer_LlenaRecordSet.Enabled = True
                                Me.Timer_LlenaRecordSet.Interval = 1000
                                
                          End If
                          timSalida.Enabled = False
                          Exit Sub
                       End If
                       RsMozo.MoveNext
                    Loop
                    MsgBox "Usuario no Encontrado", vbCritical, sMensaje
                    txtPassword.Text = ""
                    txtPassword.SetFocus
                   Exit Sub
                End If
                                
           Case Is = 2
                frmPassword.cmdOpcion.Visible = False
                frmPassword.Show vbModal
                If wEnter Then
                   txtPassword.Text = sDescrip
                End If
                
           'HUELLA
           Case Is = 3
                If lMCPV = False Then
                   If pais <> "002" Then
                        Set RsTc = Lib.OpenRecordset("SELECT * From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
                
                        If RsTc.EOF Then
                         nTC = 0
                        Else
                         nTC = IIf(IsNull(RsTc!nVenta), 0, IIf(IsNull(RsTc!nVenta), 0, RsTc!nVenta))
                        End If

                        wInicio = False
                        If nTC = 0 Then
                         MsgBox "Error: No se ha ingresado el Tipo de Cambio", vbCritical, sMensaje
                         Exit Sub
                        End If
                    End If
                End If
                wEnterHuella = False
                frmVerificacionHuella.Show vbModal
                
                If wEnterHuella Then
                    If lMCPV Then
                        If sModulo = "INFOREST" Then
                           sUsuario = sVar1
                           Unload Me
                        Else
                           'audirotia
                           registroAccesoAuditoria "I", sVar1

                           Me.Timer_LlenaRecordSet.Enabled = False
                           Me.Timer_LlenaRecordSet.Interval = 0
                           frmCargoMozo.Show vbModal
                           Me.Timer_LlenaRecordSet.Enabled = True
                           Me.Timer_LlenaRecordSet.Interval = 1000
                        End If
                    Else
                    
                          'audirotia
                          registroAccesoAuditoria "I", sVar1
                          
                          Me.Timer_LlenaRecordSet.Enabled = False
                          Me.Timer_LlenaRecordSet.Interval = 0
                          frmCargoMozo.Show vbModal
                          Me.Timer_LlenaRecordSet.Enabled = True
                          Me.Timer_LlenaRecordSet.Interval = 1000
                             
                             
                    End If
                    timSalida.Enabled = False
                    Exit Sub
                End If
           Case Is = 1
                If lHARDkey Then
                    '----------Verifica Llave HK----------------------------------
                    If hk.ValidaLlave Then
                        'MsgBox "Fallo la validacion de la llave", vbCritical, "Aviso"
                        Dim Result As Boolean
                        If sModulo = "INFOREST" Then
                            Result = hk.FinalizarConexion(Aplicacion.PuntoVenta)
                        End If
                        If sModulo = "ADICION" Then
                            Result = hk.FinalizarConexion(Aplicacion.Adicion)
                        End If
                        End
                    End If
                    '--------------------------------------------------------------
                End If
                End

    End Select
    txtPassword.SelStart = Len(txtPassword.Text)
    txtPassword.SetFocus
End Sub

'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long



Private Sub Form_Activate()
On Error GoTo fin
    txtPassword.SetFocus
fin:
End Sub

Private Sub Form_Load()
   Centrar Me
   
    Me.Width = 12000
    Me.Height = 9000
   
   inicio = True
   timSalida.Enabled = False
   timSalida.Interval = nSalir
   lblHora.Caption = Format(FechaServidor(), "HH:mm:ss")


    ListView1.Visible = False
    Imagemensaje.Visible = False
    Timer_girar.Enabled = False
    frmocualtamensaje.Visible = False


   lblconexion.Caption = sRuta & " : " & sMDB
   'Frame3.Visible = False
               
   If lSiab Then
      cmdConsultaSaldo.Visible = True
      cmdOpcion(1).Height = 830
   End If
   txtPassword.Text = ""
   If lMCPV Then
      Isql = "select tCodigoUsuario as Codigo, tResumido, tPassword as tValor, tBandaMagnetica, nTamano=0, nValor=0, thuella from vGrupousuario Where lActivo = 1 And lModulo01 = 1"
   Else
      Isql = "select * from vMozo where lActivo = 1 Order by nBoton"
   End If
   
   Set RsMozo = Lib.OpenRecordset(Isql, Cn)
'   txtFecha.Caption = Format(FechaServidor(), "dddd, mmm d yyyy") '"long date")
    txtFecha.Caption = " - " & Format(FechaServidor(), "dd/mm/yyyy")

    'GCAA 26032020
   Isql = "select isnull(trutalogo,'') as tRutaLogo, isnull(trutaimadia,'') as tRutaImaDia from tparametro"
    Set rsparam = Lib.OpenRecordset(Isql, Cn)
    



    'On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(App.Path & "\bmps\Cliente.avi") Then
       ImgLogo.Visible = False
       imgLogoPic.Visible = True
       With mmControl
            .FileName = App.Path & "\bmps\Cliente.avi"
            .Command = "open"
            .hWndDisplay = imgLogoPic.hwnd
            .DeviceType = "AVIVideo"
            .From = 0
            .Notify = True
            .Command = "play"
       End With
    Else
       If fso.FileExists(rsparam!trutalogo) Then '(App.Path & "\bmps\ima\ImaLogo.jpg") Then
          ImgLogo.Picture = LoadPicture(rsparam!trutalogo) '(App.Path & "\bmps\ima\ImaLogo.jpg")
        Else
            If fso.FileExists(App.Path & "\bmps\ima\ImaLogo.jpg") Then
                ImgLogo.Picture = LoadPicture(App.Path & "\bmps\ima\ImaLogo.jpg")
            Else
                If fso.FileExists(rsparam!trutalogo) Then
                    ImgLogo.Picture = LoadPicture(App.Path & "\bmps\Logo.jpg")
                End If
            End If
       End If
    End If

    If fso.FileExists(rsparam!trutaimadia) Then '(App.Path & "\bmps\ima\ImaDia.jpg") Then
          Image3.Picture = LoadPicture(rsparam!trutaimadia) '(App.Path & "\bmps\ima\ImaDia.jpg")
    Else
        If fso.FileExists(App.Path & "\bmps\ima\ImaSinDia.jpg") Then
            Image3.Picture = LoadPicture(App.Path & "\bmps\ima\ImaSinDia.jpg")
        Else
            If fso.FileExists(App.Path & "\bmps\Cliente.jpg") Then
                Image3.Picture = LoadPicture(App.Path & "\bmps\Cliente.jpg")
            Else
            
            End If
            'Image3.Picture = LoadPicture(App.Path & "\bmps\Cliente.jpg")
        End If
    End If
    
    
    If fso.FileExists(App.Path & "\bmps\logo.jpg") Then
       ImgLogo2.Picture = LoadPicture(App.Path & "\bmps\Logo.jpg")
    End If
        
    If fso.FileExists(App.Path & "\bmps\Paises\" & pais & ".jpg") Then
           ImagePais.Picture = LoadPicture(App.Path & "\bmps\Paises\" & pais & ".jpg")
    End If
'    Select Case pais
'        Case Is = "001"
'            If fso.FileExists(App.Path & "\bmps\Paises\001.jpg") Then
'               ImagePais.Picture = LoadPicture(App.Path & "\bmps\Paises\001.jpg")
'            End If
'        Case Is = "002"
'            If fso.FileExists(App.Path & "\bmps\Paises\002.jpg") Then
'               ImagePais.Picture = LoadPicture(App.Path & "\bmps\Paises\002.jpg")
'            End If
'        Case Else
'            If fso.FileExists(App.Path & "\bmps\Paises\000.jpg") Then
'               ImagePais.Picture = LoadPicture(App.Path & "\bmps\Paises\000.jpg")
'            End If
'    End Select
    
'    If fso.FileExists(App.Path & "\bmps\Pais.jpg") Then
'       ImagePais.Picture = LoadPicture(App.Path & "\bmps\Pais.jpg")
'    End If
    
    
    Set fso = Nothing
    'txtPassword.SetFocus
    Call Color
End Sub
Private Sub val_huella(Index As Integer)
        Select Case Index
           Case Is = 0 ' Aceptar
                                               
                If lHARDkey Then
                    ValidacionEntrarConLicencia
                End If
                If txtPassword.Text = "" Then
                   MsgBox "Ingrese su password", vbExclamation, sMensaje
                   txtPassword.SetFocus
                   Exit Sub
                Else
                    RsMozo.MoveFirst
                    Do While Not RsMozo.EOF
                       If Desencapsula(IIf(IsNull(RsMozo!tvalor), "", RsMozo!tvalor)) = UCase(txtPassword.Text) Or (IIf(IsNull(RsMozo!tBandaMagnetica), "", RsMozo!tBandaMagnetica) = Encapsula(UCase(Extrae(txtPassword.Text))) And Encapsula(UCase(Extrae(txtPassword.Text))) <> "") Then
                             sVar1 = RsMozo!tResumido
                             If RsMozo!nTamano = 1 Then
                                Dim sCambio As String
                                MsgBox "Por motivo de seguridad, deberá cambiar su password", vbCritical, sMensaje
                                
                                frmPassword.Caption = "Ingrese su nueva clave"
                                frmPassword.Show vbModal
                                If Not wEnter Then
                                   Exit Sub
                                End If
                                sCambio = sDescrip
                                frmPassword.Caption = "Confirme su nueva clave"
                                frmPassword.Show vbModal
                                If Not wEnter Then
                                   Exit Sub
                                End If
                                If sCambio <> sDescrip Or sDescrip = "" Then
                                   MsgBox "Confirmación erronea, no se realizó el cambio", vbCritical, sMensaje
                                   Exit Sub
                                End If
                                RsMozo!nTamano = 0
                                RsMozo!tvalor = Encapsula(sDescrip)
                             End If
                          sPassword = UCase(txtPassword.Text)
                          sMozo = RsMozo!codigo
                          lSomelier = IIf(IsNull(RsMozo!nValor), 0, RsMozo!nValor)
                          txtPassword.Text = ""
                          If lMCPV Then
                             If sModulo = "INFOREST" Then
                                sUsuario = sVar1
                                Unload Me
                             Else
                             
                                'audirotia
                                
                                registroAccesoAuditoria "I", sVar1
                                
                                'auditoria


                                Me.Timer_LlenaRecordSet.Enabled = False
                                Me.Timer_LlenaRecordSet.Interval = 0
                                frmCargoMozo.Show vbModal
                                Me.Timer_LlenaRecordSet.Enabled = True
                                Me.Timer_LlenaRecordSet.Interval = 1000
                                
                                
                             End If
                          Else
                          
                                'audirotia
                                
                                registroAccesoAuditoria "I", sVar1
                                
                                'auditoria
                                
                                
                                Me.Timer_LlenaRecordSet.Enabled = False
                                Me.Timer_LlenaRecordSet.Interval = 0
                                frmCargoMozo.Show vbModal
                                Me.Timer_LlenaRecordSet.Enabled = True
                                Me.Timer_LlenaRecordSet.Interval = 1000
                             
                             
                          End If
                          timSalida.Enabled = False
                          Exit Sub
                       End If
                       RsMozo.MoveNext
                    Loop
                    MsgBox "Usuario no Encontrado", vbCritical, sMensaje
                    txtPassword.Text = ""
                    txtPassword.SetFocus
                   Exit Sub
                End If
                                
           Case Is = 2
                frmPassword.cmdOpcion.Visible = False
                frmPassword.Show vbModal
                If wEnter Then
                   txtPassword.Text = sDescrip
                   If Trim(sDescrip) <> "" Then
                    cmdOpcion_Click (0)
                   End If
                End If
                
           'HUELLA
           Case Is = 3
                               
                wEnterHuella = False
                frmVerificarHuellaMarcacion.Show vbModal
                
                If wEnterHuella Then
                                                
                    timSalida.Enabled = False
                    Exit Sub
                End If
                
                
           Case Is = 1
                If lHARDkey Then
                    '----------Verifica Llave HK----------------------------------
                    If hk.ValidaLlave Then
                        'MsgBox "Fallo la validacion de la llave", vbCritical, "Aviso"
                        Dim Result As Boolean
                        If sModulo = "INFOREST" Then
                            Result = hk.FinalizarConexion(Aplicacion.PuntoVenta)
                        End If
                        If sModulo = "ADICION" Then
                            Result = hk.FinalizarConexion(Aplicacion.Adicion)
                        End If
                        End
                    End If
                    '--------------------------------------------------------------
                End If
                End
                End Select
    
End Sub





Private Sub ImageHelp_Click()
frmAbout.Show vbModal
End Sub

'Private Sub imgInforest_Click(Index As Integer)
'
'    CommonDialogcolor.ShowColor
'    frmcontenido.backColor = CommonDialogcolor.Color
'    ListView1.backColor = CommonDialogcolor.Color
'    Text2.backColor = CommonDialogcolor.Color
'
'    On Error Resume Next
'    Open App.Path & "\ColorFondo.txt" For Output As #1
'    Print #1, CommonDialogcolor.Color
'    Close #1
'
'End Sub
Private Sub Color()
    On Error Resume Next
    Dim strLinea As String
    Open App.Path & "\ColorFondo.txt" For Input As #1
    Line Input #1, strLinea
    Close #1
    
    If strLinea <> "" Then
        frmcontenido.backColor = strLinea
        ListView1.backColor = strLinea
        Text2.backColor = strLinea
    End If
    
End Sub



Private Sub imgingresar_Click()
    If Trim(txtPassword.Text) <> "" Then
        cmdOpcion_Click (0)
    End If
End Sub

Private Sub imgNewProceso_Click(Index As Integer)
    'MsgBox "Proceso en Desarrollo, Aun no culminado!!!", vbInformation, sMensaje
    marcacion = 1
    val_huella (3)
End Sub
Private Sub ValidaVersiones()
On Error GoTo fin
    
'     Dim CnSeg As New Connection
'
'    Set CnSeg = New Connection
'    CnSeg.Provider = "SQLOLEDB"
'    CnSeg.CursorLocation = adUseServer
'    CnSeg.ConnectionString = "User ID=" & sUserName & _
'                          ";password=" & sPassword & _
'                          ";Data Source=" & sRuta & _
'                          ";DataTypeCompatibility=80" & _
'                          ";Initial Catalog=INFSEGURIDAD"
'    CnSeg.CommandTimeout = 250
'    CnSeg.Open
    If CnSeg.State = 0 Then
        CnSeg.Open
    End If
    Dim VersionInfoRest, VersionFEInfoHotel, VersionActualFEInfoHotel As String
    
   ' VersionInfoRest = Calcular("select top 1 VersionInfoRest as Codigo from TPARAMETROVERSION ", CnSeg)
    Open App.Path & "\version.txt" For Input As #1   ' Abre el archivo para recibir los datos.
    Line Input #1, VersionInfoRest
    Close #1


    
    Dim sVersionExe As String
    sVersionExe = App.Major & "." & App.Minor & "." & App.Revision
    
'    If VersionFEInfoHotel <> VersionActualFEInfoHotel Then
'        Label2(2).Caption = "Motor FE " & VersionFEInfoHotel & " Click Aqui"
'        Label2(2).Visible = True
'        Frame4.Visible = True
'    Else
'        If (VersionInfoHotel <> sVersionExe) Then
'            Label2(1).Visible = True
'            Label2(1).Caption = "InfoHotel " & VersionInfoHotel & " Click Aqui"
'            Frame4.Visible = True
'        End If
'    End If

    If (VersionInfoRest <> sVersionExe) Then
        LblVer.Visible = True
        LblVer.Caption = "InfoRest " & VersionInfoRest & " "
        Frame4.Visible = True
    End If
Exit Sub
fin:
End Sub
Private Sub imgOpcion_Click(Index As Integer)
    Select Case Index
           Case Is = 0 ' Aceptar
                'TIPO CAMBIO
                If lMCPV = False Then
                    If pais <> "002" Then
                        Dim RsTc As New ADODB.Recordset
                        Set RsTc = Lib.OpenRecordset("SELECT * From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
                
                        If RsTc.EOF Then
                         nTC = 0
                        Else
                         nTC = IIf(IsNull(RsTc!nVenta), 0, IIf(IsNull(RsTc!nVenta), 0, RsTc!nVenta))
                        End If
                
                        wInicio = False
                        If nTC = 0 Then
                         MsgBox "Error: No se ha ingresado el Tipo de Cambio", vbCritical, sMensaje
                         txtPassword.Text = ""
                         txtPassword.SetFocus
                         Exit Sub
                        End If
                    End If
                End If
                
                If lHARDkey Then
                    ValidacionEntrarConLicencia
                End If
                If txtPassword.Text = "" Then
                   MsgBox "Ingrese su password", vbExclamation, sMensaje
                   txtPassword.SetFocus
                   Exit Sub
                Else
                    RsMozo.MoveFirst
                    Do While Not RsMozo.EOF
                       If Desencapsula(IIf(IsNull(RsMozo!tvalor), "", RsMozo!tvalor)) = UCase(txtPassword.Text) Or (IIf(IsNull(RsMozo!tBandaMagnetica), "", RsMozo!tBandaMagnetica) = Encapsula(UCase(Extrae(txtPassword.Text))) And Encapsula(UCase(Extrae(txtPassword.Text))) <> "") Then
                             sVar1 = RsMozo!tResumido
                             If RsMozo!nTamano = 1 Then
                                Dim sCambio As String
                                MsgBox "Por motivo de seguridad, deberá cambiar su password", vbCritical, sMensaje
                                
                                frmPassword.Caption = "Ingrese su nueva clave"
                                frmPassword.Show vbModal
                                If Not wEnter Then
                                   Exit Sub
                                End If
                                sCambio = sDescrip
                                frmPassword.Caption = "Confirme su nueva clave"
                                frmPassword.Show vbModal
                                If Not wEnter Then
                                   Exit Sub
                                End If
                                If sCambio <> sDescrip Or sDescrip = "" Then
                                   MsgBox "Confirmación erronea, no se realizó el cambio", vbCritical, sMensaje
                                   Exit Sub
                                End If
                                RsMozo!nTamano = 0
                                RsMozo!tvalor = Encapsula(sDescrip)
                             End If
                          sPassword = UCase(txtPassword.Text)
                          sMozo = RsMozo!codigo
                          lSomelier = IIf(IsNull(RsMozo!nValor), 0, RsMozo!nValor)
                          txtPassword.Text = ""
                          If lMCPV Then
                             If sModulo = "INFOREST" Then
                                sUsuario = sVar1
                                Unload Me
                             Else
                             
                                'audirotia
                                
                                registroAccesoAuditoria "I", sVar1
                                
                                'auditoria


                                Me.Timer_LlenaRecordSet.Enabled = False
                                Me.Timer_LlenaRecordSet.Interval = 0
                                frmCargoMozo.Show vbModal
                                Me.Timer_LlenaRecordSet.Enabled = True
                                Me.Timer_LlenaRecordSet.Interval = 1000
                                
                                
                             End If
                          Else
                          
                                'audirotia
                                
                                registroAccesoAuditoria "I", sVar1
                                
                                'auditoria
                                
                                
                                Me.Timer_LlenaRecordSet.Enabled = False
                                Me.Timer_LlenaRecordSet.Interval = 0
                                frmCargoMozo.Show vbModal
                                Me.Timer_LlenaRecordSet.Enabled = True
                                Me.Timer_LlenaRecordSet.Interval = 1000
                             
                             
                          End If
                          timSalida.Enabled = False
                          Exit Sub
                       End If
                       RsMozo.MoveNext
                    Loop
                    MsgBox "Usuario no Encontrado", vbCritical, sMensaje
                    txtPassword.Text = ""
                    txtPassword.SetFocus
                   Exit Sub
                End If
                                
           Case Is = 2
                frmPassword.cmdOpcion.Visible = False
                frmPassword.Show vbModal
                If wEnter Then
                   txtPassword.Text = sDescrip
                   If Trim(sDescrip) <> "" Then
                    cmdOpcion_Click (0)
                   End If
                End If
                
           'HUELLA
           Case Is = 3
                If lMCPV = False Then
                    If pais <> "002" Then
                        Set RsTc = Lib.OpenRecordset("SELECT * From TTIPOCAMBIO WHERE (fFecha = {fn CURDATE() })", Cn)
                
                        If RsTc.EOF Then
                         nTC = 0
                        Else
                         nTC = IIf(IsNull(RsTc!nVenta), 0, IIf(IsNull(RsTc!nVenta), 0, RsTc!nVenta))
                        End If
                        
                        wInicio = False
                        If nTC = 0 Then
                         MsgBox "Error: No se ha ingresado el Tipo de Cambio", vbCritical, sMensaje
                         Exit Sub
                        End If
                    End If
                End If
                
                wEnterHuella = False
                frmVerificacionHuella.Show vbModal
               
                    If wEnterHuella Then
                        If lMCPV Then
                            If sModulo = "INFOREST" Then
                               sUsuario = sVar1
                               Unload Me
                            Else
                               'audirotia
                               registroAccesoAuditoria "I", sVar1
    
                               Me.Timer_LlenaRecordSet.Enabled = False
                               Me.Timer_LlenaRecordSet.Interval = 0
                               frmCargoMozo.Show vbModal
                               Me.Timer_LlenaRecordSet.Enabled = True
                               Me.Timer_LlenaRecordSet.Interval = 1000
                            End If
                        Else
                        
                              'audirotia
                              registroAccesoAuditoria "I", sVar1
                              
                              Me.Timer_LlenaRecordSet.Enabled = False
                              Me.Timer_LlenaRecordSet.Interval = 0
                              frmCargoMozo.Show vbModal
                              Me.Timer_LlenaRecordSet.Enabled = True
                              Me.Timer_LlenaRecordSet.Interval = 1000
                                 
                                 
                        End If
                        timSalida.Enabled = False
                        Exit Sub
    '                                Me.Timer_LlenaRecordSet.Enabled = False
    '                                Me.Timer_LlenaRecordSet.Interval = 0
    '                                frmCargoMozo.Show vbModal
    '                                Me.Timer_LlenaRecordSet.Enabled = True
    '                                Me.Timer_LlenaRecordSet.Interval = 1000
                    End If
                 
           Case Is = 1
                If lHARDkey Then
                    '----------Verifica Llave HK----------------------------------
                    If hk.ValidaLlave Then
                        'MsgBox "Fallo la validacion de la llave", vbCritical, "Aviso"
                        Dim Result As Boolean
                        If sModulo = "INFOREST" Then
                            Result = hk.FinalizarConexion(Aplicacion.PuntoVenta)
                        End If
                        If sModulo = "ADICION" Then
                            Result = hk.FinalizarConexion(Aplicacion.Adicion)
                        End If
                        End
                    End If
                    '--------------------------------------------------------------
                End If
                End

    End Select
    txtPassword.SelStart = Len(txtPassword.Text)
    'txtPassword.SetFocus ' Retirado por causar error - 28/10/2024 FERM
End Sub





Private Sub mmControl_Done(NotifyCode As Integer)
     With mmControl
       .From = 0
       .Command = "play"
    End With
End Sub





Private Sub Timer_girar_Timer()
   Generar_Movimiento
End Sub

Private Sub Timer_LlenaRecordSet_Timer()
    ListView1.ListItems.Clear
     Llenar_RecordSet
End Sub

Private Sub Timer2_Timer()
On Error GoTo fin
If ActivoActualizador Then
    Call ValidaVersiones
End If
fin:
End Sub

Private Sub txtPassword_Click()
        frmPassword.cmdOpcion.Visible = False
        frmPassword.Show vbModal
        If wEnter Then
           txtPassword.Text = sDescrip
           If Trim(sDescrip) <> "" Then
            cmdOpcion_Click (0)
           End If
        End If

End Sub

Private Sub txtPassword_GotFocus()
'    If Trim(txtPassword.Text) <> "" Then
'      'cmdOpcion_Click (0)
'   End If

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdOpcion_Click (0)
   End If
End Sub

Private Sub timSalida_Timer()
   Do While Screen.ActiveForm.Name <> "frmMozoUsuario"
      Unload Screen.ActiveForm
   Loop
End Sub

Public Sub ReseteaTimer()
   frmMozoUsuario.timSalida.Enabled = False
   frmMozoUsuario.timSalida.Enabled = True
End Sub

Private Sub Reloj()
   Static last_time As Date

   Dim Cx As Single
   Dim cy As Single
   Dim num As Single
   Dim radius As Single
   Dim theta As Single

     If last_time = Now Then Exit Sub

     last_time = Now
     Picture1.Cls
     Picture1.ForeColor = vbBlue
     Cx = Picture1.ScaleWidth / 2
     cy = Picture1.ScaleHeight / 2

     ' Horas
     num = 5 * (DatePart("h", last_time) + DatePart("n", last_time) / _
                                 60 + DatePart("s", last_time) / 3600)
     theta = MinutesToRadians(num)
     radius = Picture1.ScaleWidth * 0.24
     Picture1.ForeColor = &H26050F
     Picture1.DrawWidth = 7
     Picture1.Line (Cx, cy)-Step(radius * Cos(theta), -radius * Sin(theta))

     ' Los Minutos
     num = DatePart("n", last_time)
     theta = MinutesToRadians(num)
     radius = Picture1.ScaleWidth * 0.37
     Picture1.ForeColor = &H26050F
     Picture1.DrawWidth = 6
     Picture1.Line (Cx, cy)-Step(radius * Cos(theta), -radius * Sin(theta))

     ' Los segundos
     num = DatePart("s", last_time)
     theta = MinutesToRadians(num)
     radius = Picture1.ScaleWidth * 0.34
     Picture1.ForeColor = &H1E32CD
     Picture1.DrawWidth = 4
     Picture1.Line (Cx, cy)-Step(radius * Cos(theta), -radius * Sin(theta))
    'Call RetornarMensajes
    
 End Sub
 Private Function MinutesToRadians(ByVal num As Single) As Single
     MinutesToRadians = (15 - num) * 2 * PI / 60
 End Function
Private Sub Timer1_Timer()
    lblHora.Caption = Format(FechaServidor(), "HH:mm:ss")
    'Reloj
    
End Sub

Private Sub Llenar_RecordSet()
    Dim X As Integer
    Dim Item As ListItem
    Isql = "usp_listadoMensajes"
    Set rsMensajeCocina = Lib.OpenRecordset(Isql, Cn)
    If rsMensajeCocina.EOF Or rsMensajeCocina.BOF Then
       'Frame3.Visible = False
       ListView1.Visible = False
       Imagemensaje.Visible = False
       Timer_girar.Enabled = False
       frmocualtamensaje.Visible = False
    Else
            With ListView1
                .View = lvwReport
                .ListItems.Clear
                .ColumnHeaders.Clear
            End With
            nroElementos = rsMensajeCocina.RecordCount
            Timer_LlenaRecordSet.Interval = (nroElementos + 1) * 1000
            Me.MousePointer = vbHourglass
            ListView1.ColumnHeaders.Add , , "", 7000
            
            rsMensajeCocina.MoveFirst
            ListView1.ListItems.Clear
            While Not rsMensajeCocina.EOF
                Set Item = ListView1.ListItems.Add(, , rsMensajeCocina.Fields(0))
                rsMensajeCocina.MoveNext
            Wend
            If rsMensajeCocina.EOF Then
                ListView1.ListItems.Add , , "-"
            End If
            Me.MousePointer = vbDefault
            'Frame3.Visible = True
            Imagemensaje.Visible = True
            ListView1.Visible = True
            Timer_girar.Enabled = True
            frmocualtamensaje.Visible = True
            Timer_girar.Interval = 2000
    End If
End Sub
Private Sub Generar_Movimiento()
    Dim temp As String
    Dim X As Integer
    Dim numero_lista As Integer
    numero_lista = ListView1.ListItems.Count
    temp = ListView1.ListItems.Item(1)
    For X = 1 To numero_lista - 1
        ListView1.ListItems.Item(X) = ListView1.ListItems.Item(X + 1)
        If X = numero_lista - 1 Then
            ListView1.ListItems.Item(numero_lista) = temp
        End If
    Next X
End Sub





