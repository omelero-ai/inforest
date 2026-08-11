VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOptimizar 
   Caption         =   "Optimizar Base de Datos"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   8100
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Estado "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      Left            =   120
      TabIndex        =   10
      Top             =   200
      Width           =   2355
      Begin VB.Image imgRojo 
         Height          =   2655
         Left            =   120
         Picture         =   "FrmOptimizar.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2055
      End
      Begin VB.Image imgVerde 
         Height          =   2655
         Left            =   120
         Picture         =   "FrmOptimizar.frx":C2A0
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "InfoRest"
      Height          =   1455
      Left            =   2520
      TabIndex        =   7
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton Command1 
         Caption         =   "Optimizar BD - INFOREST"
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
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label LblUltimoInfo 
         Alignment       =   2  'Center
         Caption         =   "-"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Bizlinks"
      Enabled         =   0   'False
      Height          =   1815
      Left            =   2520
      TabIndex        =   1
      Top             =   1680
      Width           =   5535
      Begin VB.CommandButton Command2 
         Caption         =   "Optimizar BD - BIZLINKS"
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
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   5055
      End
      Begin VB.TextBox TxtDias 
         Height          =   285
         Left            =   2520
         TabIndex        =   2
         Text            =   "7"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "dias en los temporales"
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   270
         Width           =   1665
      End
      Begin VB.Label LblUltimoBiz 
         Alignment       =   2  'Center
         Caption         =   "-"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label Label4 
         Caption         =   "Solo dejar DOC de los ultimos"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   270
         Width           =   2175
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   3480
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13759
            MinWidth        =   5292
         EndProperty
      EndProperty
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
End
Attribute VB_Name = "FrmOptimizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
imgRojo.Visible = True
imgVerde.Visible = False


StatusBar.Panels(1).Text = "Optimizando BD...."
Dim oComandoCabecera As clsComando

Set oComandoCabecera = New clsComando
If Not oComandoCabecera.CreateCmdSp("sp_OptimizarBD", Cn) Then
     Set oComandoCabecera = Nothing
End If

oComandoCabecera.CreateParameter "@xbd", adVarChar, adParamInputOutput, 500, sMDB
If Not oComandoCabecera.GetParamOK Then
     Set oComandoCabecera = Nothing
End If

If Not oComandoCabecera.ExecSP Then
     Set oComandoCabecera = Nothing
End If



  
 If oComandoCabecera.GetParameterValue("@xbd") = "Optimizado" Then
    StatusBar.Panels(1).Text = "Terminado...."
    imgRojo.Visible = False
    imgVerde.Visible = True
    
    Isql = "insert into LOG_OPTIMIZACION(Usuario, fregistro, BD) " & _
                "values ( '" & sUsuario & "' , getdate() ,'" & sMDB & "')"
    
    Cn.Execute Isql
    
    
    MsgBox "Proceso realizado satisfactoriamente", vbInformation, sMensaje
 Else
    imgRojo.Visible = True
    imgVerde.Visible = False
    StatusBar.Panels(1).Text = oComandoCabecera.GetParameterValue("@xbd") & " ...."
    MsgBox oComandoCabecera.GetParameterValue("@xbd"), vbCritical, sMensaje
End If
Ultimo

End Sub

Sub Ultimo()
   Dim UltimoInfo, UltimoBiz As String
   
   UltimoInfo = Calcular("select TOP 1 'Ultima Optimización: ' + CONVERT(varchar(100), fregistro) + ' / Usuario: ' + usuario as codigo  from LOG_optimizacion where BD= '" & sMDB & "'  order by fregistro desc", Cn)
   
   If UltimoInfo = "0" Then
   LblUltimoInfo.Caption = "-"
   Else
   LblUltimoInfo.Caption = UltimoInfo
   End If
   
    UltimoBiz = Calcular("select TOP 1 'Ultima Optimización: ' + CONVERT(varchar(100), fregistro) + ' / Usuario: ' + usuario as codigo  from LOG_optimizacion where BD= 'BIZLINKS_PROD'  order by fregistro desc", Cn)
   If UltimoBiz = "0" Then
   LblUltimoBiz.Caption = "-"
   Else
   LblUltimoBiz.Caption = UltimoBiz
   End If
End Sub

Private Sub Command2_Click()
    imgRojo.Visible = True
    imgVerde.Visible = False
    
    
    StatusBar.Panels(1).Text = "Optimizando BD...."
    Dim oComandoCabecera As clsComando
    
    Set oComandoCabecera = New clsComando
    If Not oComandoCabecera.CreateCmdSp("sp_OptimizarBD_Bizlinks", Cn) Then
         Set oComandoCabecera = Nothing
    End If
    
    oComandoCabecera.CreateParameter "@_BD_INFOREST", adVarChar, adParamInputOutput, 500, sMDB
    oComandoCabecera.CreateParameter "@_dias", adVarChar, adParamInput, 2, TxtDias.Text
    
    If Not oComandoCabecera.GetParamOK Then
         Set oComandoCabecera = Nothing
    End If
    
    If Not oComandoCabecera.ExecSP Then
         Set oComandoCabecera = Nothing
    End If
    


  
 If oComandoCabecera.GetParameterValue("@_BD_INFOREST") = "Optimizado" Then
    StatusBar.Panels(1).Text = "Terminado...."
    imgRojo.Visible = False
    imgVerde.Visible = True
    
    Isql = "insert into LOG_OPTIMIZACION(Usuario, fregistro, BD) " & _
            "values ( '" & sUsuario & "' , getdate() ,'BIZLINKS_PROD')"
    
    Cn.Execute Isql
    MsgBox "Proceso realizado satisfactoriamente", vbInformation, sMensaje
 Else
    imgRojo.Visible = True
    imgVerde.Visible = False
    StatusBar.Panels(1).Text = oComandoCabecera.GetParameterValue("@_BD_INFOREST") & " ...."
    MsgBox oComandoCabecera.GetParameterValue("@_BD_INFOREST"), vbCritical, sMensaje
End If
Ultimo
End Sub

Private Sub Form_Load()
     imgRojo.Visible = True
   imgVerde.Visible = False
   
    If (lFEBiz) Then
        Frame3.Enabled = True
    End If
   
  Ultimo
End Sub

