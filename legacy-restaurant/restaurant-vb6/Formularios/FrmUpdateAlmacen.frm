VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmUpdateAlmacen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Migrar Datos"
   ClientHeight    =   5445
   ClientLeft      =   3465
   ClientTop       =   2775
   ClientWidth     =   7350
   Icon            =   "FrmUpdateAlmacen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7350
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   810
      Left            =   0
      ScaleHeight     =   750
      ScaleWidth      =   7290
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4635
      Width           =   7350
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Index           =   0
         Left            =   4680
         Picture         =   "FrmUpdateAlmacen.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1095
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
         Height          =   520
         Index           =   1
         Left            =   6120
         Picture         =   "FrmUpdateAlmacen.frx":13CC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Este es un proceso crítico del sistema. No realice ninguna operación mientras no termine el mismo."
         ForeColor       =   &H00800000&
         Height          =   720
         Left            =   1290
         TabIndex        =   5
         Top             =   60
         Width           =   2805
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Cuidado :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   15
         TabIndex        =   4
         Top             =   15
         Width           =   1230
      End
   End
   Begin MSComCtl2.Animation aniVideo 
      Height          =   540
      Left            =   1560
      TabIndex        =   0
      Top             =   2850
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   953
      _Version        =   393216
      FullWidth       =   49
      FullHeight      =   36
   End
   Begin VB.Frame Frame1 
      Height          =   4125
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   7335
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Pasos para la Actualización"
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   555
         TabIndex        =   14
         Top             =   0
         Width           =   1950
      End
      Begin VB.Label lblProceso 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Verificando Conectividad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Index           =   0
         Left            =   2640
         TabIndex        =   13
         Top             =   645
         Width           =   2250
      End
      Begin VB.Label lblProceso 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actualizando Tablas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   1
         Left            =   2640
         TabIndex        =   12
         Top             =   1125
         Width           =   1845
      End
      Begin VB.Label lblProceso 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finalizando la Actualización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Index           =   2
         Left            =   2640
         TabIndex        =   11
         Top             =   1620
         Width           =   2475
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   0
         Left            =   2040
         Picture         =   "FrmUpdateAlmacen.frx":14BE
         Top             =   645
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   1
         Left            =   2040
         Picture         =   "FrmUpdateAlmacen.frx":16D1
         Top             =   1125
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   2
         Left            =   2040
         Picture         =   "FrmUpdateAlmacen.frx":18E4
         Top             =   1605
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Servidor Remoto :"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   300
         TabIndex        =   10
         Top             =   3525
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Base Datos Remota :"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   3840
         Width           =   1515
      End
      Begin VB.Label lblBddCentral 
         AutoSize        =   -1  'True
         Caption         =   "Base Datos Central :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1710
         TabIndex        =   8
         Top             =   3525
         Width           =   1455
      End
      Begin VB.Label lblBddLocal 
         AutoSize        =   -1  'True
         Caption         =   "Base Datos Central :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1710
         TabIndex        =   7
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404000&
         X1              =   0
         X2              =   7320
         Y1              =   3405
         Y2              =   3405
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actualización desde Servidor Remoto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   75
      TabIndex        =   15
      Top             =   90
      Width           =   3930
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   480
      Left            =   0
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "FrmUpdateAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsTablas As Recordset
Dim RsTemp As Recordset

Sub Inicializa()
    lblProceso(0).ForeColor = &H808080
    lblProceso(1).ForeColor = &H808080
    lblProceso(2).ForeColor = &H808080
    imgProceso(0).Visible = False
    imgProceso(1).Visible = False
    imgProceso(2).Visible = False
    
    aniVideo.AutoPlay = True
    aniVideo.Visible = True
End Sub
 
Function VerificaConexion() As Boolean
On Error GoTo err

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
            VerificaConexion = True
        Else
            VerificaConexion = False
        End If
    End If
    Exit Function
    
err:
    VerificaConexion = False
End Function

Private Sub cmdOpcion_Click(Index As Integer)
'On Error GoTo LeerError
    If Index = 0 Then
        If MsgBox("¿Seguro que desea continuar con la actualización de los datos?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
            If VerificaConexion Then
                Me.cmdOpcion(0).Enabled = False
                Screen.MousePointer = vbHourglass
                Inicializa
                lblProceso(0).ForeColor = &H80000012
                DoEvents
                    
                imgProceso(0).Visible = True
                If lAlmacenRemoto = True Then
                    CnAlmacenRemoto.Execute "sp_ActualizaReceta"
                    lblProceso(1).ForeColor = &H80000012
                        
                    CargaTablasAlmacenRemoto
                    imgProceso(1).Visible = True
                End If
                DoEvents
                
                lblProceso(2).ForeColor = &H80000012
                DoEvents
                
                imgProceso(2).Visible = True
                DoEvents
                aniVideo.AutoPlay = False
                aniVideo.Visible = False
            End If
        End If
        Me.cmdOpcion(0).Enabled = True
        Screen.MousePointer = vbDefault
        Set CnAlmacenRemoto = Nothing
    Else
        Unload Me
    End If
'    Exit Sub
'LeerError:
'    DoEvents
'    aniVideo.AutoPlay = False
'    aniVideo.Visible = False
'    Me.cmdOpcion(0).Enabled = True
'    Screen.MousePointer = vbDefault
'    MsgBox err.Description, vbCritical, sMensaje
End Sub

Private Sub Form_Load()
    Dim nFrames As Long
    Centrar Me
    lblBddCentral.Caption = UCase(sRutaAlmacenRemoto)
    lblBddLocal.Caption = UCase(sMDBAlmacenRemoto)
    On Error Resume Next
    aniVideo.Open App.Path & "\bmps\FileMove.avi"
End Sub

