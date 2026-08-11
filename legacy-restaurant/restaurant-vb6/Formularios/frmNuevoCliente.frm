VERSION 5.00
Begin VB.Form frmNuevoCliente 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Cliente"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10845
   Icon            =   "frmNuevoCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAgenteRetencion 
      BackColor       =   &H00E2DEDB&
      Caption         =   "Es Agente de retención?"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Frame frmDatosSeparados 
      BackColor       =   &H00E2DEDB&
      Caption         =   "Datos separados"
      Enabled         =   0   'False
      Height          =   2655
      Left            =   5280
      TabIndex        =   27
      Top             =   0
      Width           =   5535
      Begin VB.TextBox txtNombre 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   50
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   330
         Width           =   3840
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
         Caption         =   "Nombre"
         Height          =   555
         Index           =   10
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox txtSegundoNombre 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   50
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   900
         Width           =   3840
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
         Caption         =   "Segundo Nombre"
         Height          =   555
         Index           =   11
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   810
         Width           =   1275
      End
      Begin VB.TextBox txtApellido 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   50
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1500
         Width           =   3840
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
         Caption         =   "Apellido"
         Height          =   555
         Index           =   12
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1410
         Width           =   1275
      End
      Begin VB.TextBox txtSegundoApellido 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   200
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2100
         Width           =   3840
      End
      Begin VB.CommandButton cmdOpcion 
         BackColor       =   &H00F2FAED&
         Caption         =   "Segundo Apellido"
         Height          =   555
         Index           =   13
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2010
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdValidaDNI 
      BackColor       =   &H00E7ECD5&
      Caption         =   "Verificar DNI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Caption         =   "Tipo Responsable"
      Height          =   555
      Index           =   9
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3480
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txtTipoResponsable 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   40
      MaxLength       =   200
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.TextBox txtUbigeo 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2560
      MaxLength       =   200
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1320
   End
   Begin VB.CommandButton cmdUbigeo 
      BackColor       =   &H00F2FAED&
      Caption         =   "Ubigeo"
      Height          =   555
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4680
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Caption         =   "&Urbanizacion"
      Height          =   555
      Index           =   8
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1750
      Width           =   1275
   End
   Begin VB.TextBox txtUrbanizacion 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      MaxLength       =   200
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3840
   End
   Begin VB.CommandButton cmdValidar 
      BackColor       =   &H00E7ECD5&
      Caption         =   "Validar Ruc SUNAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdTipoCliente 
      BackColor       =   &H00F2FAED&
      Caption         =   "Tipo Cliente"
      Height          =   555
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4065
      Width           =   1275
   End
   Begin VB.TextBox txtTipoCliente 
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
      Height          =   375
      Left            =   45
      MaxLength       =   200
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4185
      Width           =   3840
   End
   Begin VB.TextBox txtEnlace 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   40
      MaxLength       =   200
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3600
      Width           =   3840
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Enlace"
      Height          =   555
      Index           =   7
      Left            =   3960
      TabIndex        =   13
      Top             =   3480
      Width           =   1275
   End
   Begin VB.TextBox txtTipoIdentidad 
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
      Height          =   375
      Left            =   45
      MaxLength       =   15
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   3810
   End
   Begin VB.CommandButton cmTipoIdentidad 
      BackColor       =   &H00F2FAED&
      Caption         =   "Tipo Identidad"
      Height          =   555
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   30
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Caption         =   "&Correo"
      Height          =   555
      Index           =   6
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2895
      Width           =   1275
   End
   Begin VB.TextBox txtCorreo 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   40
      MaxLength       =   200
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3000
      Width           =   3840
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00E7ECD5&
      Caption         =   "Mostrar Visor de Precios"
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
      Index           =   5
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5295
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00E2DEDB&
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
      Index           =   4
      Left            =   2610
      Picture         =   "frmNuevoCliente.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5310
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00E2DEDB&
      Caption         =   "Aceptar"
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
      Index           =   3
      Left            =   3960
      Picture         =   "frmNuevoCliente.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5310
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Caption         =   "&Dirección"
      Height          =   555
      Index           =   2
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2340
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Caption         =   "Razón &Social"
      Height          =   555
      Index           =   1
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1185
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00F2FAED&
      Caption         =   "&Identificador"
      Height          =   555
      Index           =   0
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1275
   End
   Begin VB.TextBox txtRuc 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      MaxLength       =   15
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   3840
   End
   Begin VB.TextBox txtRazonSocial 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      MaxLength       =   200
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1250
      Width           =   3840
   End
   Begin VB.TextBox txtDireccion 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      MaxLength       =   200
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2400
      Width           =   3840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label lblcondicion 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   6000
      Width           =   2535
   End
End
Attribute VB_Name = "frmNuevoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTipoIdentidad As ADODB.Recordset
Dim RsTipoCliente As ADODB.Recordset
Dim RsTipoResponsable As ADODB.Recordset
Dim RsDatosClienteFac As ADODB.Recordset

Dim sCodigoTipoIdentidad As String
Dim sCodigoTipoCliente As String
Dim sCodigoTipoResponsable As String

Dim validaTipoIdentidad As Boolean

Dim lDatosSeparados As Boolean


'cambio validar DNI
'Dim Isql As String
Dim RsParametroDNI As Recordset
Private Sub cmdUbigeo_Click()
   Dim xCriterio As String
   xCriterio = sCodigo
   Isql = "Select tCodigo as Codigo, tDistrito as Descripcion, tProvincia as Provincia, tDepartamento as Departamento from TUBIGEO order by tCodigo asc"
   
   frmBusca.TipoOperacion = "UBIGEO"
   frmBusca.cboCriterio.Enabled = True
   frmBusca.nPredeterm = 1

   Call ConfGrilla(4, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                          "Distrito", 2, "Descripcion", 1500, 0, 0, "", _
                                          "Provincia", 2, "Provincia", 2500, 0, 0, "", _
                                          "Departamento", 2, "Departamento", 3000, 0, 0, "")
   frmBusca.Show vbModal
   
   If Not wEnter Then
      sCodigo = xCriterio
      Exit Sub
   End If
   
   txtUbigeo.Text = sCodigo
   sCodigo = xCriterio
End Sub

Public Function ValidaIdentidad(ByVal parametros As String) As String
On Error GoTo fin

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(App.path & "\ConsultaInfomatica\ConsultaInfomatica.exe") Then
    
        Dim wsh As Object
        Set wsh = CreateObject("WScript.Shell")
        
        Dim Cmd As String
        Cmd = """" & App.path & "\ConsultaInfomatica\ConsultaInfomatica.exe"" """ & parametros & """"

        Dim proc As Object
        Set proc = wsh.Exec(Cmd)
    
        Dim resultado As String
        resultado = ""
        
        Do While Not proc.StdOut.AtEndOfStream
            resultado = resultado & proc.StdOut.ReadLine
        Loop
        
        ValidaIdentidad = resultado
        
        Exit Function
        
    End If

    ValidaIdentidad = ""
    Exit Function
fin:
    MsgBox "Error número: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description, vbCritical, "Error detectado"
           
    ValidaIdentidad = ""
End Function


Private Sub cmdValidaDNI_Click()

    On Error GoTo fin
    Screen.MousePointer = vbHourglass
    Dim xDat As String
    Dim xxDat As String
    
    Dim xRazSoc As String, xEst As String, xCon As String, xDir As String
    Dim xRazSocX As Long, xEstX As Long, xConX As Long, xDirX As Long
    Dim xRazSocY As Long, xEstY As Long, xConY As Long, xDirY As Long
    Dim xWml As New XMLHTTP
    Dim xxWml As New XMLHTTP
    
    Me.TxtNombre = ""
    
    If Not IsNumeric(Trim(TxtRUC.Text)) Then
        MsgBox "formato incorrectO, Verificar!!!", vbInformation, sMensaje
        Exit Sub
    End If
    If Len(Trim(TxtRUC.Text)) <> 8 Then
        MsgBox "longitud incorrecto, Verificar!!!", vbInformation, sMensaje
        Exit Sub
    End If
    
    
    
    'MOTOR CONSULTA
    
    Dim resultado As String
    Dim envio As String
    
    envio = "01," & TxtRUC.Text & ","
    resultado = ValidaIdentidad(envio)
    
    If Not resultado = "" Then
    
        Dim obj As Object
        Dim json As Object
        Dim sc As Object
        
        Set sc = CreateObject("MSScriptControl.ScriptControl")
        sc.Language = "JScript"
        sc.AddCode "function parse(j){return eval('(' + j + ')');}"
        sc.AddCode "function getCI(o,k){k=k.toLowerCase();for(var p in o){if(p.toLowerCase()==k)return o[p];}return '';}"
        
        Set obj = sc.Run("parse", resultado)
        
        Dim Estado As String
        Estado = sc.Run("getCI", obj, "status")
        
        If Estado = "1" Then
            Me.txtRazonSocial.Text = sc.Run("getCI", obj, "apellidos_nombres")
            Me.TxtNombre.Text = sc.Run("getCI", obj, "apellidos_nombres")
            Me.txtDireccion.Text = sc.Run("getCI", obj, "direccion")
            Me.txtUbigeo.Text = sc.Run("getCI", obj, "ubigeo")
            
            MsgBox sc.Run("getCI", obj, "msg")
            
            Exit Sub 'TERMINA PROCESO
            
        End If
        
    End If
    
    'FIN
    

    xWml.Open "POST", "https://cloudservices.infomatica.pe/api/consultadni/" & Trim(TxtRUC.Text), False
    xWml.send
    
    If xWml.Status = 200 Then
        xDat = Trim(xWml.responseText)
        xDat = Replace(Replace(Replace(xDat, "," + Chr(34), vbNewLine), "{", ""), "}", "")
        xDat = Replace(xDat, Chr(34), "")
        
        Dim MATRIZ() As String
        MATRIZ = Split(xDat, vbNewLine)
        
        MATRIZ(0) = Replace(MATRIZ(0), "status:", "")
        If (MATRIZ(0) = 1) Then
            MATRIZ(2) = Replace(MATRIZ(2), "dni:", "")
            MATRIZ(3) = Replace(MATRIZ(3), "apellidos_nombres:", "")
            MATRIZ(4) = Replace(MATRIZ(4), "direccion:", "")
            MATRIZ(5) = Replace(MATRIZ(5), "ubigeo:", "")
            
            'Me.txtUrbanizacion.Text = Calcular("select tdistrito as codigo from tubigeo where tcodigo='" & MATRIZ(5) & "'", Cn)
            Me.txtRazonSocial.Text = MATRIZ(3)
            
             If lDatosSeparados = True Then
                Me.TxtNombre.Text = MATRIZ(3)
            End If
            'txtRazonSocial.Text = MATRIZ(0)
            Me.txtDireccion.Text = MATRIZ(4)
            'Me.txtUbigeo = MATRIZ(5)
        
        Else
           ' Call RucSUNAT(Trim(txtRuc.Text))
        End If
    ElseIf xWml.Status = 500 Then
        MsgBox "No hay resultados a la consulta", vbInformation, "JNE"
    ElseIf xWml.Status = 12007 Then
        MsgBox "Error en la conexion con Internet", vbInformation, sMensaje
    Else
        MsgBox "Error en el proceso" & xWml.Status, vbInformation, sMensaje
    End If
    
    Set xWml = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
fin:
    MsgBox "Mensaje: " & Error
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdValidar_Click()
    'Call RucSUNAT(Trim(txtRuc.Text))
    
    On Error GoTo fin
    Screen.MousePointer = vbHourglass
    Dim xDat As String
    Dim xxDat As String
    
    Dim xRazSoc As String, xEst As String, xCon As String, xDir As String
    Dim xRazSocX As Long, xEstX As Long, xConX As Long, xDirX As Long
    Dim xRazSocY As Long, xEstY As Long, xConY As Long, xDirY As Long
    Dim xWml As New XMLHTTP
    Dim xxWml As New XMLHTTP
    
    Me.TxtNombre = ""
    
    If Not IsNumeric(Trim(TxtRUC.Text)) Then
        MsgBox "formato incorrectO, Verificar!!!", vbInformation, sMensaje
        Exit Sub
    End If
    If Len(Trim(TxtRUC.Text)) <> 11 Then
        MsgBox "longitud incorrecto, Verificar!!!", vbInformation, sMensaje
        Exit Sub
    End If
    
    
    
    
    'MOTOR CONSULTA
    
    Dim resultado As String
    Dim envio As String
    
    envio = "02," & TxtRUC.Text & ","
    resultado = ValidaIdentidad(envio)
    
    If Not resultado = "" Then
    
        Dim obj As Object
        Dim json As Object
        Dim sc As Object
        
        Set sc = CreateObject("MSScriptControl.ScriptControl")
        sc.Language = "JScript"
        
        sc.AddCode "function parse(j){return eval('(' + j + ')');}"
        sc.AddCode "function getCI(o,k){k=k.toLowerCase();for(var p in o){if(p.toLowerCase()==k)return o[p];}return '';}"

        Set obj = sc.Run("parse", resultado)
        
        Dim Estado As String
        Estado = sc.Run("getCI", obj, "status")
        
        If Estado = "1" Then
        
            Me.txtRazonSocial.Text = sc.Run("getCI", obj, "razonsocial")
            Me.TxtNombre.Text = sc.Run("getCI", obj, "razonsocial")
            Me.txtDireccion.Text = sc.Run("getCI", obj, "direccion")
            Me.txtUbigeo.Text = sc.Run("getCI", obj, "ubigeo")
            
            MsgBox sc.Run("getCI", obj, "msg")
            
            'RETENCIONES
            
            If UCase(xtipodoc) = "F" Then
            If xTotal >= val(tMontoRetencion) Then
    
                '9016
                xxWml.Open "GET", RutaConsultaRucRetenciones & Trim(TxtRUC.Text), False
                xxWml.send
                
                If xxWml.Status = 200 Then
                    xxDat = Trim(xxWml.responseText)
                    xxDat = Replace(Replace(Replace(xxDat, "," + Chr(34), vbNewLine), "{", ""), "}", "")
                    xxDat = Replace(xxDat, Chr(34), "")
                    
                     Dim xMATRIZ_TEMP() As String
                     xMATRIZ_TEMP = Split(xxDat, vbNewLine)
                    
                    xMATRIZ_TEMP(0) = Replace(xMATRIZ_TEMP(3), "status:", "")
                    If (xMATRIZ_TEMP(0) <> "Ruc:null") Then
                        Cn.Execute "update TCLIENTE set AgenteRetencion=1  where tcodigocliente='" & sCodigo & "'"
                        chkAgenteRetencion.value = 1
                    Else
                         Cn.Execute "update TCLIENTE set AgenteRetencion=0  where tcodigocliente='" & sCodigo & "'"
                         chkAgenteRetencion.value = 0
                    End If
                ElseIf xxWml.Status = 500 Then
                    MsgBox "No hay resultados a la consulta", vbInformation, "JNE"
                ElseIf xxWml.Status = 12007 Then
                    MsgBox "Error en la conexion con Internet", vbInformation, sMensaje
                Else
                    MsgBox "Error en el proceso" & xWml.Status, vbInformation, sMensaje
                End If
            
                Set xxWml = Nothing
                End If
            End If
            Set xWml = Nothing
            Screen.MousePointer = vbDefault
            
            Exit Sub 'TERMINA PROCESO
            
        End If
        
    End If
    
    ' FIN
    
    
    
    xWml.Open "POST", RutaConsultaRuc & "/" & Trim(TxtRUC.Text), False
    xWml.send
    
    If xWml.Status = 200 Then
        xDat = Trim(xWml.responseText)
        xDat = Replace(Replace(Replace(xDat, "," + Chr(34), vbNewLine), "{", ""), "}", "")
        xDat = Replace(xDat, Chr(34), "")
        
        Dim MATRIZ() As String
        MATRIZ = Split(xDat, vbNewLine)
        
        MATRIZ(0) = Replace(MATRIZ(0), "status:", "")
        If (MATRIZ(0) = 1) Then
            MATRIZ(2) = Replace(MATRIZ(2), "ruc:", "")
            MATRIZ(3) = Replace(MATRIZ(3), "razonSocial:", "")
            MATRIZ(4) = Replace(MATRIZ(4), "direccion:", "")
            MATRIZ(5) = Replace(MATRIZ(5), "ubigeo:", "")
            
            Me.txtUrbanizacion.Text = Calcular("select tdistrito as codigo from tubigeo where tcodigo='" & MATRIZ(5) & "'", Cn)
            Me.txtRazonSocial.Text = MATRIZ(3)
            
            If lDatosSeparados = True Then
                Me.TxtNombre.Text = MATRIZ(3)
            End If
            'txtRazonSocial.Text = MATRIZ(0)
            Me.txtDireccion.Text = MATRIZ(4)
            Me.txtUbigeo = MATRIZ(5)
        
        Else
           ' Call RucSUNAT(Trim(txtRuc.Text))
        End If
    ElseIf xWml.Status = 500 Then
        MsgBox "No hay resultados a la consulta", vbInformation, "JNE"
    ElseIf xWml.Status = 12007 Then
        MsgBox "Error en la conexion con Internet", vbInformation, sMensaje
    Else
        MsgBox "Error en el proceso" & xWml.Status, vbInformation, sMensaje
    End If
    
    If UCase(xtipodoc) = "F" Then
            If xTotal >= val(tMontoRetencion) Then
    
                '9016
                xxWml.Open "GET", RutaConsultaRucRetenciones & Trim(TxtRUC.Text), False
                xxWml.send
                
                If xxWml.Status = 200 Then
                    xxDat = Trim(xxWml.responseText)
                    xxDat = Replace(Replace(Replace(xxDat, "," + Chr(34), vbNewLine), "{", ""), "}", "")
                    xxDat = Replace(xxDat, Chr(34), "")
                    
                     Dim xMATRIZ() As String
                     xMATRIZ = Split(xxDat, vbNewLine)
                    
                    xMATRIZ(0) = Replace(xMATRIZ(3), "status:", "")
                    If (xMATRIZ(0) <> "Ruc:null") Then
                        Cn.Execute "update TCLIENTE set AgenteRetencion=1  where tcodigocliente='" & sCodigo & "'"
                        chkAgenteRetencion.value = 1
                    Else
                         Cn.Execute "update TCLIENTE set AgenteRetencion=0  where tcodigocliente='" & sCodigo & "'"
                         chkAgenteRetencion.value = 0
                    End If
                ElseIf xxWml.Status = 500 Then
                    MsgBox "No hay resultados a la consulta", vbInformation, "JNE"
                ElseIf xxWml.Status = 12007 Then
                    MsgBox "Error en la conexion con Internet", vbInformation, sMensaje
                Else
                    MsgBox "Error en el proceso" & xWml.Status, vbInformation, sMensaje
                End If
            
                Set xxWml = Nothing
        End If
    End If
    Set xWml = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
fin:
    'Call RucSUNAT(Trim(txtRuc.Text))
    MsgBox "Mensaje: " & Error
    Screen.MousePointer = vbDefault
    
End Sub



Private Sub txtRuc_Change()
' validar dni
    Me.TxtRUC.Text = Replace(Me.TxtRUC, " ", "")
End Sub

Private Sub txtRuc_LostFocus()
    If Trim(TxtRUC) = "" Then
     Exit Sub
    End If
End Sub
Private Sub cmdOpcion_Click(Index As Integer)

    Dim xtTipoIdentidad As String
   Dim Numero As Boolean
   Dim te As String

   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   Select Case Index
          Case Is = 0 'Ruc
               frmKeyBoard.tipo = "Cliente"
               frmKeyBoard.txtResultado = TxtRUC.Text
               frmKeyBoard.Show vbModal
               
               If Trim(sDescrip) = "" Then
                Exit Sub
               End If
               
                If pais <> "002" Then 'PERU - BOLIVIA

'                    Dim RsValidacionDatos As Recordset
'                    Set RsValidacionDatos = Lib.OpenRecordset("exec usp_ValidaTipoIdentidad '" & sCodigoTipoIdentidad & "','" & sDescrip & "'", Cn)
'                    If RsValidacionDatos.RecordCount > 0 Then
'                        RsValidacionDatos.MoveFirst
'                        If Left(RsValidacionDatos!Respuesta, 1) = "1" Then
'                            MsgBox "Validacion identificador: " & RsValidacionDatos!Respuesta, vbInformation
'                            txtRuc.Text = ""
'                            wEnter = False
'                            Exit Sub
'                        End If
'                    End If
'                    If Trim(txtTipoIdentidad.Text) = "RUC" Then
'                        If Not ValidaRuc(sDescrip) And pais = "000" Then
'                           MsgBox "El número RUC ingresado no es válido", vbCritical, sMensaje
'                           'txtRuc.Text = ""
'                           wEnter = False
'                           Exit Sub
'                        End If
'                    End If
                    
                Else 'ECUADOR
                
                    'If chkpasaporte.value = 0 Then
'                        If Calcular("Select isnull(nValor,0) As Codigo from vtipoidentidad where Codigo= '" & sCodigoTipoIdentidad & "'", Cn) Then
'
'                            If Len(Trim(sDescrip)) = 13 Or Len(Trim(sDescrip)) = 10 Then
'                                If xlTipoDocumento = True Then
'                                    If Not ValidaEcuadorCedulaRuc(sDescrip) Then
'                                       MsgBox "Identificador no Válido", vbCritical, sMensaje
'                                       wEnter = False
'                                       Exit Sub
'                                    End If
'                                End If
'                            Else
'                               MsgBox "La longitud del Identificador debe ser 10(Cédula) ó 13(RUC)", vbCritical, sMensaje
'                               wEnter = False
'                               Exit Sub
'                            End If
'
'                        End If
                    'End If
                    
                End If
                
                If Calcular("Select isnull(nValor,0) As Codigo from vtipoidentidad where Codigo= '" & sCodigoTipoIdentidad & "'", Cn) Then
                    Dim RsValidacionDatos As Recordset
                    Set RsValidacionDatos = Lib.OpenRecordset("exec usp_ValidaTipoIdentidad '" & sCodigoTipoIdentidad & "','" & sDescrip & "'", Cn)
                    If RsValidacionDatos.RecordCount > 0 Then
                        RsValidacionDatos.MoveFirst
                        If Left(RsValidacionDatos!respuesta, 1) = "1" Then
                            MsgBox "Validacion identificador: " & RsValidacionDatos!respuesta, vbInformation
                            TxtRUC.Text = ""
                            wEnter = False
                            Exit Sub
                        End If
                    End If
                End If

                If Trim(txtTipoIdentidad.Text) = "RUC" And pais = "000" Then
                    If Not ValidaRuc(sDescrip) Then
                       MsgBox "El número RUC ingresado no es válido", vbCritical, sMensaje
                       'txtRuc.Text = ""
                       wEnter = False
                       Exit Sub
                    End If
                End If
                    
                If val(Calcular("select tIdentidad as Codigo from TCLIENTE where tIdentidad = '" & sDescrip & "'", Cn)) > 0 Then
                   MsgBox "Identificador Repetido", vbCritical, sMensaje
                   wEnter = False
                   Exit Sub
                End If
                xlTipoDocumento = False

                TxtRUC.Text = IIf(wEnter, sDescrip, TxtRUC.Text)
          
          Case Is = 1 ' Razon social
               frmKeyBoard.txtResultado = txtRazonSocial.Text
               frmKeyBoard.Show vbModal
               txtRazonSocial.Text = IIf(wEnter, sDescrip, txtRazonSocial.Text)
          
          Case Is = 2 ' Direccion
               frmKeyBoard.txtResultado = txtDireccion.Text
               frmKeyBoard.Show vbModal
               txtDireccion.Text = IIf(wEnter, sDescrip, txtDireccion.Text)
               
          Case Is = 3 ' Aceptar
               Dim nCorrela As String
               
               ' cambios para validar DNI
               If RsParametroDNI!lValidaDNI = True Then
               If txtTipoIdentidad.Text = "DNI" Then
               Numero = modProcedimiento.ValidarDNI(LTrim(TxtRUC.Text))
                    If Numero = False Then
                    MsgBox "El DNI ingresado no es valido", vbCritical, sMensaje
                    Exit Sub
                    End If
               End If
               End If
               '---------------------------------
               'Chequea Datos
               If TxtRUC.Text = "" Then MsgBox "No se ha ingresado el Identificador", vbExclamation, sMensaje: Exit Sub
               If Calcular(" select isnull(lActivaDatosSeparados,0) as codigo from vTipoIdentidad where codigo='" & sCodigoTipoIdentidad & "' ", Cn) Then
                    If Trim(Me.TxtNombre.Text) = "" Then MsgBox "Ingrese Nombre de cliente", vbExclamation, sMensaje: Exit Sub
                    If Trim(Me.txtSegundoNombre.Text) = "" Then MsgBox "Ingrese segundo nombre de cliente", vbExclamation, sMensaje: Exit Sub
                    If Trim(Me.txtApellido.Text) = "" Then MsgBox "Ingrese apellido de cliente", vbExclamation, sMensaje: Exit Sub
                    If Trim(Me.txtSegundoApellido.Text) = "" Then MsgBox "Ingrese segundo apellido de cliente", vbExclamation, sMensaje: Exit Sub
                    Me.txtRazonSocial.Text = Trim(Me.TxtNombre.Text) & " " & Trim(Me.txtSegundoNombre.Text) & " " & Trim(Me.txtApellido.Text) & " " & Trim(Me.txtSegundoApellido.Text)
               Else
                    If txtRazonSocial = "" Then MsgBox "Ingrese la Razón Social", vbExclamation, sMensaje: Exit Sub
               End If
               
'               If txtRazonSocial = "" Then MsgBox "Ingrese la Razón Social", vbExclamation, sMensaje: Exit Sub
               If txtTipoIdentidad.Text = "" Then MsgBox "Seleccione Tipo de Identidad", vbExclamation, sMensaje: Exit Sub
               If txtTipoIdentidad.Text = "Seleccionar --->" Then MsgBox "Seleccione Tipo de Identidad", vbExclamation, sMensaje: Exit Sub
               
               If sCodigoTipoIdentidad = "" Then MsgBox "Seleccione Tipo de Identidad", vbExclamation, sMensaje: Exit Sub
               
               If lSAP Then
                    If txtTipoCliente.Text = "" Then MsgBox "Seleccione Tipo de Cliente", vbExclamation, sMensaje: Exit Sub
                    If txtTipoCliente.Text = "Seleccionar --->" Then MsgBox "Seleccione Tipo de Cliente", vbExclamation, sMensaje: Exit Sub
                    If sCodigoTipoCliente = "" Then MsgBox "Seleccione Tipo de Cliente", vbExclamation, sMensaje: Exit Sub
               End If
               
               If txtCorreo.Text <> "" Then
                    If Not Validar_Email(txtCorreo.Text) Then
                        MsgBox "El Correo ingresado no es válido", vbCritical, sMensaje
                        wEnter = False
                        Exit Sub
                    End If
               End If
               If lFEpape Then
                    If sCodigoTipoIdentidad = "02" Or sCodigoTipoIdentidad = "01" Then
                        If txtCorreo.Text <> "" Then
                             If Not Validar_Email(txtCorreo.Text) Then
                                 MsgBox "El Correo ingresado no es válido", vbCritical, sMensaje
                                 wEnter = False
                                 Exit Sub
                             End If
                        Else
                            MsgBox "Ingrese un correo Electronico", vbCritical, sMensaje
                        End If
                    End If
               End If

               'Validacion de Identificador
               If pais = "003" Then ' ARGENTINA
               
                If sCodigoTipoResponsable = "" Then MsgBox "Seleccione Tipo responsable", vbExclamation, sMensaje: Exit Sub
                If Trim(Me.txtDireccion) = "" Then MsgBox "Ingrese direcciòn", vbExclamation, sMensaje: Exit Sub
                    validaTipoIdentidad = False
                    validaTipoIdentidad = Calcular("select isnull(nvalor,0) as codigo from vTipoidentidad where Codigo='" & sCodigoTipoIdentidad & "' ", Cn)
                    If validaTipoIdentidad = True Then
                         If lLongitud And Len(Trim(TxtRUC.Text)) <> nLongitud Then
                            MsgBox "La longitud del Identificador debe ser " & nLongitud, vbCritical, sMensaje
                            wEnter = False
                            Exit Sub
                         ElseIf Not lLongitud And Len(Trim(TxtRUC.Text)) < nLongitud Then
                            MsgBox "La longitud del Identificador debe ser mayor igual a " & nLongitud, vbCritical, sMensaje
                            wEnter = False
                            Exit Sub
                         End If
                         Dim ValTemp As String
                         ValTemp = ValidaCUIT(TxtRUC.Text)
                         If ValTemp = "" Then
                            'MsgBox "El número ingresado no es válido", vbCritical, sMensaje
                            wEnter = False
                            Exit Sub
                          Else
                            TxtRUC.Text = ValTemp
                          End If
                         xtTipoIdentidad = ""
                    End If
               ElseIf pais = "002" Then ' ECUADOR
'                    If sCodigoTipoIdentidad = "01" Or sCodigoTipoIdentidad = "02" Then
'                        If Len(Trim(txtRuc.Text)) = 13 Or Len(Trim(txtRuc.Text)) = 10 Then
'                             If xlTipoDocumento = True Then
'                                 If Not ValidaEcuadorCedulaRuc(txtRuc.Text) Then
'                                    MsgBox "Identificador no Válido", vbCritical, sMensaje
'                                    wEnter = False
'                                    Exit Sub
'                                 End If
'                             End If
'                         Else
'                            MsgBox "La longitud del Identificador debe ser 10(Cédula) ó 13(RUC)", vbCritical, sMensaje
'                            wEnter = False
'                            Exit Sub
'                         End If
'
'                        'SEGUN SRI
'                        If Len(Trim(txtRuc.Text)) = 10 Then
'                           xtTipoIdentidad = "01"
'                        ElseIf Len(Trim(txtRuc.Text)) = 13 Then
'                           xtTipoIdentidad = "02"
'                        End If
'
'                    Else
'                        xtTipoIdentidad = "03"
'                        sCodigoTipoIdentidad = "03"
'                    End If
               Else 'PERU - BOLIVIA
'                    Dim RsValidacionDatos2 As Recordset
'                    Set RsValidacionDatos2 = Lib.OpenRecordset("exec  usp_ValidaTipoIdentidad '" & sCodigoTipoIdentidad & "','" & txtRuc.Text & "'", Cn)
'                    If RsValidacionDatos2.RecordCount > 0 Then
'                        RsValidacionDatos2.MoveFirst
'                        If Left(RsValidacionDatos2!Respuesta, 1) = "1" Then
'                            MsgBox "Validacion identificador : " & RsValidacionDatos2!Respuesta, vbInformation
'                            Exit Sub
'                        End If
'                    End If
'                    If Trim(txtTipoIdentidad.Text) = "RUC" Then
'                        If Not ValidaRuc(txtRuc.Text) And pais = "000" Then
'                           MsgBox "El número RUC ingresado no es válido", vbCritical, sMensaje
'                           'txtRuc.Text = ""
'                           wEnter = False
'                           Exit Sub
'                        End If
'                    End If
                    If Trim(txtTipoIdentidad.Text) = "DNI" Then
                        If Len(Trim(TxtRUC.Text)) <> 8 And pais = "000" Then
                           MsgBox "El número DNI ingresado no tiene la Longitud correcta", vbCritical, sMensaje
                           'txtRuc.Text = ""
                           wEnter = False
                           Exit Sub
                        End If
                    End If
                    xtTipoIdentidad = ""
               End If
               
                If Calcular("Select isnull(nValor,0) As Codigo from vtipoidentidad where Codigo= '" & sCodigoTipoIdentidad & "'", Cn) Then
                    Dim RsValidacionDatos2 As Recordset
                    Set RsValidacionDatos2 = Lib.OpenRecordset("exec  usp_ValidaTipoIdentidad '" & sCodigoTipoIdentidad & "','" & TxtRUC.Text & "'", Cn)
                    If RsValidacionDatos2.RecordCount > 0 Then
                        RsValidacionDatos2.MoveFirst
                        If Left(RsValidacionDatos2!respuesta, 1) = "1" Then
                            MsgBox "Validacion identificador : " & RsValidacionDatos2!respuesta, vbInformation
                            Exit Sub
                        End If
                    End If
                    If Trim(txtTipoIdentidad.Text) = "RUC" And pais = "000" Then
                        If Not ValidaRuc(TxtRUC.Text) Then
                           MsgBox "El número RUC ingresado no es válido", vbCritical, sMensaje
                           'txtRuc.Text = ""
                           wEnter = False
                           Exit Sub
                        End If
                    End If
                
                End If
               
            
               If frmBusquedaRapida.wAdiciona Then
                  If val(Calcular("select tIdentidad as Codigo from TCLIENTE where tIdentidad = '" & Apostrofe_v2(TxtRUC.Text) & "'", Cn)) > 0 Then
                     MsgBox "Identificador Repetido", vbCritical, sMensaje
                     wEnter = False
                     Exit Sub
                  End If
               
                  'Obtiene el Correlativo
                  nCorrela = Calcular("select Max(tCodigoCliente) as Codigo from TCLIENTE", Cn)
                
                  If IsNull(nCorrela) Or nCorrela = "" Then
                     sCodigo = "00001"
                  Else
'                     sCodigo = Lib.Correlativo(nCorrela, 5)
                     'AJMM 27-09-2025
                    Dim letras As String
                    Dim numeroCor As Long
                    Dim i As Integer
                    
                    letras = ""
                    numeroCor = 0
                    
                    ' Extraemos las letras del inicio
                    i = 1
                    Do While i <= Len(nCorrela) And IsLetter(Mid$(nCorrela, i, 1))
                        letras = letras & Mid$(nCorrela, i, 1)
                        i = i + 1
                    Loop
                    
                    ' Extraemos la parte numérica
                    If i <= Len(nCorrela) Then
                        numeroCor = CLng(Mid$(nCorrela, i))
                    End If
                    
                    ' Incrementamos número
                    numeroCor = numeroCor + 1
                    
                    ' Cantidad de dígitos que deben quedar para número
                    Dim numDig As Integer
                    numDig = 5 - Len(letras)
                    
                    ' Si el número se desborda (ej: 99999 ? 100000), incrementamos letras
                    If Len(CStr(numeroCor)) > numDig Then
                        numeroCor = 1 ' reiniciamos numeración
                        letras = IncrementarLetras(letras)
                        numDig = 5 - Len(letras)
                    End If
                    
                    ' Formamos el código final (siempre 5 caracteres)
                    sCodigo = letras & Right$(String(numDig, "0") & CStr(numeroCor), numDig)
                    'sCodigo = letras & Format$(numeroCor, String$(numDig, "0"))
                    
                        
                  End If
                  
                  Isql = "insert into TCLIENTE( " & _
                         "tCodigoCliente, tEmpresa, tIdentidad, tDireccion, tCorreo, tUsuario, tTipoIdentidad, lActivo,AgenteRetencion, tEnlace, tTipoCliente, tubigeo,tUrbanizacion ,fRegistro, ttiporesponsable, tnombre, tnombre2, tApellido, tApellido2) " & _
                         "values ('" & sCodigo & "', " & _
                                 " '" & Apostrofe_v2(txtRazonSocial.Text) & "', " & _
                                 " '" & Apostrofe_v2(TxtRUC.Text) & "', " & _
                                 " '" & Apostrofe_v2(txtDireccion.Text) & "', " & _
                                 " '" & Apostrofe_v2(txtCorreo.Text) & "', " & _
                                 " '" & sUsuario & "', " & _
                                 " '" & sCodigoTipoIdentidad & "', " & _
                                                          1 & ", " & _
                                 " '" & chkAgenteRetencion.value & "'," & _
                                 " '" & txtEnlace.Text & "', " & _
                                 " '" & sCodigoTipoCliente & "', " & _
                                 " '" & Me.txtUbigeo.Text & "', " & _
                                 " '" & Me.txtUrbanizacion.Text & "', " & _
                                 " getdate(), '" & sCodigoTipoResponsable & "', '" & Trim(Me.TxtNombre.Text) & "', '" & Trim(Me.txtSegundoNombre.Text) & "', '" & Trim(Me.txtApellido.Text) & "', '" & Trim(Me.txtSegundoApellido.Text) & "' )"
               Else

                  If val(Calcular("select tIdentidad as Codigo from TCLIENTE where tIdentidad = '" & TxtRUC.Text & "' and tCodigoCliente<>'" & sCodigo & "' ", Cn)) > 0 Then
                     MsgBox "Identificador Repetido", vbCritical, sMensaje
                     wEnter = False
                     Exit Sub
                  End If
                                 
                  Isql = "Update TCLIENTE  SET " & _
                         "tIdentidad='" & Apostrofe_v2(TxtRUC.Text) & "', " & _
                         "tEmpresa='" & Apostrofe_v2(txtRazonSocial.Text) & "', " & _
                         "tDireccion='" & Apostrofe_v2(txtDireccion.Text) & "', " & _
                         "tCorreo='" & Apostrofe_v2(txtCorreo.Text) & "', " & _
                         "tTipoIdentidad = '" & sCodigoTipoIdentidad & "', " & _
                         "tTipoCliente = '" & sCodigoTipoCliente & "', " & _
                         "tEnlace = '" & txtEnlace.Text & "', " & _
                         "tUbigeo = '" & Me.txtUbigeo.Text & "', " & _
                         "tUrbanizacion = '" & Me.txtUrbanizacion.Text & "', " & _
                         "fRegistro=getdate() ,lreplica = 1, " & _
                         "ttiporesponsable='" & sCodigoTipoResponsable & "', " & _
                         "tNombre='" & Trim(Me.TxtNombre.Text) & "', " & _
                         "tNombre2='" & Trim(Me.txtSegundoNombre.Text) & "', " & _
                         "tApellido='" & Trim(Me.txtApellido.Text) & "', " & _
                         "tApellido2='" & Trim(Me.txtSegundoApellido.Text) & "', " & _
                          "AgenteRetencion =" & chkAgenteRetencion.value & "" & _
                         "where tCodigoCliente='" & sCodigo & "'"
               End If
               Cn.Execute Isql
               Cn.Execute "Update TCLIENTE  SET lNuevoModificado=1 where tCodigoCliente='" & sCodigo & "'"
               wEnter = True
               Unload Me
               
          Case Is = 6 ' Correo
               frmKeyBoard.txtResultado = txtCorreo.Text
               frmKeyBoard.Show vbModal
               txtCorreo.Text = IIf(wEnter, sDescrip, txtCorreo.Text)
               
          Case Is = 7 ' Enlace
               frmKeyBoard.txtResultado = txtEnlace.Text
               frmKeyBoard.Show vbModal
               txtEnlace.Text = IIf(wEnter, sDescrip, txtEnlace.Text)
               
          Case Is = 4 ' Salir
               wEnter = False
               Unload Me
               
          Case Is = 5 ' Mostrar
               Visor txtRazonSocial.Text, TxtRUC.Text, nPuerto, "N"
          Case Is = 8 ' Urbanizacion
               frmKeyBoard.txtResultado = txtUrbanizacion.Text
               frmKeyBoard.Show vbModal
               Me.txtUrbanizacion.Text = IIf(wEnter, sDescrip, txtUrbanizacion.Text)
          Case Is = 9 ' TipoResponsable


                RsTipoResponsable.MoveNext
                If RsTipoResponsable.EOF Then
                    RsTipoResponsable.MoveFirst
                End If
                sCodigoTipoResponsable = RsTipoResponsable!codigo
                'cmdTipoCliente.Caption = RsTipoCliente!tResumido
                Me.txtTipoResponsable.Text = RsTipoResponsable!descripcion
                wEnter = True
                
        Case Is = 10 ' Nombre
                frmKeyBoard.tipo = "Cliente"
               frmKeyBoard.txtResultado = Trim(Me.TxtNombre.Text)
               frmKeyBoard.Show vbModal
               Me.TxtNombre.Text = Trim(IIf(wEnter, sDescrip, TxtNombre.Text))
               Me.txtRazonSocial.Text = Trim(Me.TxtNombre.Text) & " " & Trim(Me.txtSegundoNombre.Text) & " " & Trim(Me.txtApellido.Text) & " " & Trim(Me.txtSegundoApellido.Text)
        Case Is = 11 ' Segundo Nombre
                frmKeyBoard.tipo = "Cliente"
               frmKeyBoard.txtResultado = Trim(Me.txtSegundoNombre.Text)
               frmKeyBoard.Show vbModal
               Me.txtSegundoNombre.Text = Trim(IIf(wEnter, sDescrip, txtSegundoNombre.Text))
               Me.txtRazonSocial.Text = Trim(Me.TxtNombre.Text) & " " & Trim(Me.txtSegundoNombre.Text) & " " & Trim(Me.txtApellido.Text) & " " & Trim(Me.txtSegundoApellido.Text)
        Case Is = 12 ' Apellido
               frmKeyBoard.tipo = "Cliente"
               frmKeyBoard.txtResultado = Trim(Me.txtApellido.Text)
               frmKeyBoard.Show vbModal
               Me.txtApellido.Text = Trim(IIf(wEnter, sDescrip, txtApellido.Text))
               Me.txtRazonSocial.Text = Trim(Me.TxtNombre.Text) & " " & Trim(Me.txtSegundoNombre.Text) & " " & Trim(Me.txtApellido.Text) & " " & Trim(Me.txtSegundoApellido.Text)
        Case Is = 13 ' Segundo Apellido
               frmKeyBoard.tipo = "Cliente"
               frmKeyBoard.txtResultado = Trim(Me.txtSegundoApellido.Text)
               frmKeyBoard.Show vbModal
               Me.txtSegundoApellido.Text = Trim(IIf(wEnter, sDescrip, txtSegundoApellido.Text))
               Me.txtRazonSocial.Text = Trim(Me.TxtNombre.Text) & " " & Trim(Me.txtSegundoNombre.Text) & " " & Trim(Me.txtApellido.Text) & " " & Trim(Me.txtSegundoApellido.Text)
   End Select
End Sub

Private Sub cmdTipoCliente_Click()

    RsTipoCliente.MoveNext
    If RsTipoCliente.EOF Then
        RsTipoCliente.MoveFirst
    End If
    sCodigoTipoCliente = RsTipoCliente!codigo
    cmdTipoCliente.Caption = RsTipoCliente!tResumido
    txtTipoCliente.Text = RsTipoCliente!tResumido
    wEnter = True
                              
End Sub

Private Sub cmTipoIdentidad_Click()
'    If pais <> "003" Then
'        'txtRuc.Text = ""
'    End If
    RsTipoIdentidad.MoveNext
     If RsTipoIdentidad.EOF Then
        RsTipoIdentidad.MoveFirst
     End If
     sCodigoTipoIdentidad = RsTipoIdentidad!codigo
     cmTipoIdentidad.Caption = RsTipoIdentidad!tResumido
     txtTipoIdentidad.Text = RsTipoIdentidad!tResumido
     wEnter = True
     Call SolicitaDatosSeparados
     If sCodigoTipoIdentidad = "02" Then
        chkAgenteRetencion.Visible = True
    Else
        chkAgenteRetencion.Visible = False
    End If
    
    If (Calcular("Select isnull(lActivaDatosSeparados,0) as codigo from vtipoidentidad where Codigo='" & sCodigoTipoIdentidad & "'", Cn)) Then
        lDatosSeparados = True
    Else

        lDatosSeparados = False
    End If
End Sub

Private Sub SolicitaDatosSeparados()
On Error GoTo fin
    If Calcular(" select isnull(lActivaDatosSeparados,0) as codigo from vTipoIdentidad where codigo='" & sCodigoTipoIdentidad & "' ", Cn) Then
        frmNuevoCliente.Width = 10935
        frmDatosSeparados.Enabled = True
        cmdOpcion(1).Enabled = False
    Else
        frmNuevoCliente.Width = 5415
        frmDatosSeparados.Enabled = False
        cmdOpcion(1).Enabled = True
    End If
    Exit Sub
fin:

End Sub

Private Sub Form_Initialize()
Set RsTipoIdentidad = New ADODB.Recordset
Set RsTipoCliente = New ADODB.Recordset
Set RsTipoResponsable = New ADODB.Recordset
End Sub

Private Sub Form_Load()


    
    
    ' cambio validar DNI
   Isql = "select lValidaDNI from TPARAMETRO"
   Set RsParametroDNI = Lib.OpenRecordset(Isql, Cn)
    '------------------------------------------

   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   Set RsTipoIdentidad = Lib.OpenRecordset("select tresumido,codigo from vtipoidentidad where lActivo = 1 ", Cn)
   Set RsTipoCliente = Lib.OpenRecordset("select tresumido,codigo from vtipogrupocliente where lActivo = 1", Cn)
   Set RsTipoResponsable = Lib.OpenRecordset("select tresumido,codigo,descripcion from vtiporesponsable where lActivo = 1", Cn)
   
   Select Case pais
    Case "001" 'Bolivia
        cmdOpcion(0).Caption = "&NIT"
        
    Case "003"
        cmdOpcion(9).Visible = True
        Me.txtTipoResponsable.Visible = True
        cmdOpcion(7).Visible = False
        Me.txtEnlace.Visible = False
        cmdUbigeo.Visible = False
        txtUbigeo.Visible = False
        Label1.Visible = False
    Case "002"
        cmdOpcion(0).Caption = "&Identificador"
    Case Else 'Peru, Ecuador
        cmdOpcion(0).Caption = "&Identificador"
        cmdValidaDNI.Visible = True
        cmdValidar.Visible = True

   End Select
   
   Centrar Me
   If frmBusquedaRapida.wAdiciona Then
      cmdOpcion(0).Enabled = True
      Limpiar
      If frmBusquedaRapida.nPredeterm = 2 Then
         txtRazonSocial.Text = frmBusquedaRapida.txtResultado
      Else
         TxtRUC.Text = frmBusquedaRapida.txtResultado
      End If
      txtTipoIdentidad.Text = "Seleccionar -->"
        'RsTipoIdentidad.MoveNext
        If RsTipoIdentidad.EOF Then
          RsTipoIdentidad.MoveFirst
        End If
        'sCodigoTipoIdentidad = rsTipoIdentidad!codigo
        cmTipoIdentidad.Caption = RsTipoIdentidad!tResumido
        'txtTipoIdentidad.Text = rsTipoIdentidad!tResumido
        If pais = "003" Then
            sCodigoTipoIdentidad = RsTipoIdentidad!codigo
            'cmTipoIdentidad.Caption = RsTipoIdentidad!tResumido
            txtTipoIdentidad.Text = RsTipoIdentidad!tResumido
            txtTipoCliente.Text = RsTipoIdentidad!tResumido
        Else
            sCodigoTipoIdentidad = RsTipoIdentidad!codigo
            'cmTipoIdentidad.Caption = RsTipoIdentidad!tResumido
            txtTipoIdentidad.Text = RsTipoIdentidad!tResumido
            txtTipoCliente.Text = RsTipoIdentidad!tResumido
            'txtTipoCliente.Text = "Seleccionar -->"
        End If
        
        
        If RsTipoCliente.RecordCount > 0 Then
            RsTipoCliente.MoveNext
            If RsTipoCliente.EOF Then
              RsTipoCliente.MoveFirst
            End If
            cmdTipoCliente.Caption = RsTipoCliente!tResumido
            'sCodigoTipoResponsable = RsTipoCliente!tResumido
        End If
        If RsTipoResponsable.RecordCount > 0 Then
            'RsTipoResponsable.MoveNext
            If RsTipoResponsable.EOF Then
              RsTipoResponsable.MoveFirst
            End If
            Me.txtTipoResponsable.Text = RsTipoResponsable!descripcion
            sCodigoTipoResponsable = RsTipoResponsable!codigo
        Else
            cmdOpcion(9).Enabled = False
            Me.txtTipoResponsable.Text = "No hay registros"
        End If
   Else
      cmdOpcion(0).Enabled = True
      sCodigo = IIf(frmBusquedaRapida.RsGrilla.EOF = True, "", frmBusquedaRapida.RsGrilla!codigo)
      Mostrar
   End If
   
   If lSAP Then
        txtTipoCliente.Visible = True
        cmdTipoCliente.Visible = True
   Else
        txtTipoCliente.Visible = False
        cmdTipoCliente.Visible = False
   End If
   
   
   If nPuerto > 0 Then
      cmdOpcion(5).Visible = True
   End If
   
   If Not lClub Then
    txtEnlace.Visible = False
    cmdOpcion(7).Visible = False
   End If
   Call SolicitaDatosSeparados
   
    If sCodigoTipoIdentidad <> "02" Then
        chkAgenteRetencion.Visible = False
    End If
    
    If sCodigoTipoIdentidad = "02" Then
        If lActivarAvisoRetenciones = True Then
            If xTotal >= val(tMontoRetencion) Then
                MsgBox "validar si el cliente es Agente de Retencion", vbExclamation, sMensaje
            End If
            
        End If
    End If
    
    If (Calcular("Select isnull(lActivaDatosSeparados,0) as codigo from vtipoidentidad where Codigo='" & sCodigoTipoIdentidad & "'", Cn)) Then
        lDatosSeparados = True
    Else

        lDatosSeparados = False
    End If
    
End Sub

Sub Mostrar()
    With frmBusquedaRapida.RsGrilla
        TxtRUC.Text = IIf(IsNull(!tIdentidad), "", !tIdentidad)
        txtRazonSocial = IIf(IsNull(!descripcion), "", !descripcion)
        txtDireccion = IIf(IsNull(!tDireccion), "", !tDireccion)
        txtCorreo = IIf(IsNull(!tcorreo), "", !tcorreo)
        Me.txtUrbanizacion = IIf(IsNull(!urbanizacion), "", !urbanizacion)
        Me.txtUbigeo = IIf(IsNull(!CodigoUbigeo), "", !CodigoUbigeo)
        Me.TxtNombre = Calcular("select isnull(tNombre,'') as codigo from vcliente where Codigo='" & IIf(IsNull(!codigo), "", !codigo) & "'", Cn)
        Me.txtSegundoNombre = Calcular("select isnull(tNombre2,'') as codigo from vcliente where Codigo='" & IIf(IsNull(!codigo), "", !codigo) & "'", Cn)
        Me.txtApellido = Calcular("select isnull(tApellido,'') as codigo from vcliente where Codigo='" & IIf(IsNull(!codigo), "", !codigo) & "'", Cn)
        Me.txtSegundoApellido = Calcular("select isnull(tApellido2,'') as codigo from vcliente where Codigo='" & IIf(IsNull(!codigo), "", !codigo) & "'", Cn)
        
        'GCAA 03032021
        Set RsDatosClienteFac = Lib.OpenRecordset("select * from vcliente where Codigo='" & IIf(IsNull(!codigo), "", !codigo) & "'", Cn)
        Me.chkAgenteRetencion.value = IIf(RsDatosClienteFac!AgenteRetencion = True, 1, 0)
        
        If !TipoIdentidad = "" Then
              txtTipoIdentidad.Text = "Seleccionar --->"
              sCodigoTipoIdentidad = ""
        Else
            txtTipoIdentidad = IIf(IsNull(!TipoIdentidad), "", !TipoIdentidad)
            Me.cmTipoIdentidad.Caption = IIf(IsNull(!TipoIdentidad), "", !TipoIdentidad)
            RsTipoIdentidad.MoveFirst
            RsTipoIdentidad.Find "tresumido='" & Me.cmTipoIdentidad.Caption & "'"
            If Not (RsTipoIdentidad.EOF Or RsTipoIdentidad.BOF) Then
                 sCodigoTipoIdentidad = RsTipoIdentidad!codigo
            End If
        End If
        
        If !TipoCliente = "" Then
              txtTipoCliente.Text = "Seleccionar --->"
              sCodigoTipoCliente = ""
        Else
            txtTipoCliente = IIf(IsNull(!TipoCliente), "", !TipoCliente)
            Me.cmdTipoCliente.Caption = IIf(IsNull(!TipoCliente), "", !TipoCliente)
            RsTipoCliente.MoveFirst
            RsTipoCliente.Find "tresumido='" & Me.cmdTipoCliente.Caption & "'"
            If Not (RsTipoCliente.EOF Or RsTipoCliente.BOF) Then
                 sCodigoTipoCliente = RsTipoCliente!codigo
            End If
        End If
        If !ttiporesponsable = "" Then
              Me.txtTipoResponsable.Text = "Seleccionar --->"
              sCodigoTipoResponsable = ""
        Else
            If RsTipoResponsable.RecordCount > 0 Then
                txtTipoResponsable.Text = IIf(IsNull(!DesTR), "", !DesTR)
                'Me.cmdTipoCliente.Caption = IIf(IsNull(!TipoCliente), "", !TipoCliente)
                RsTipoResponsable.MoveFirst
                RsTipoResponsable.Find "descripcion='" & Me.txtTipoResponsable.Text & "'"
                If Not (RsTipoResponsable.EOF Or RsTipoResponsable.BOF) Then
                    sCodigoTipoResponsable = RsTipoResponsable!codigo
                End If
            Else
                Me.txtTipoResponsable.Text = "No hay registros"
                sCodigoTipoResponsable = ""
                cmdOpcion(9).Enabled = False
            End If
        End If
        
        txtEnlace.Text = IIf(IsNull(!tEnlace), "", !tEnlace)
    End With
End Sub

Sub Limpiar()
  Dim Control As Object
  For Each Control In Me.Controls
        If (TypeOf Control Is TextBox) Then
            Control.Text = ""
        End If
    Next Control
End Sub

Private Sub Form_Terminate()
    Set RsTipoIdentidad = Nothing
    Set RsTipoCliente = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmNuevoCliente = Nothing
End Sub

'cambio de validar DNI
Private Sub txtRuc_GotFocus()
If txtTipoIdentidad.Text = "Seleccionar ---" Then
    MsgBox "Debe colocar un identificador"
    foco
    Exit Sub
End If
End Sub


'cambio de validar DNI
Private Function foco()
Me.txtRazonSocial.SetFocus
End Function


'cambios validar DNI
Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If RsParametroDNI!lValidaDNI = True Then
    If txtTipoIdentidad.Text = "DNI" Then
        If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
        ElseIf KeyAscii <> 8 Then
        If Not IsNumeric(Chr(KeyAscii)) Then
        Beep
        KeyAscii = 0
        End If
        End If
    End If
End If
End Sub
' Función IsLetter
Private Function IsLetter(ByVal char As String) As Boolean
    IsLetter = (Asc(char) >= 65 And Asc(char) <= 90) Or (Asc(char) >= 97 And Asc(char) <= 122)
End Function
Private Function IncrementarLetras(ByVal letras As String) As String
    Dim i As Integer
    Dim arr() As Byte
    Dim resultado As String
    
    If letras = "" Then
        resultado = "A"
    Else
        arr = StrConv(letras, vbFromUnicode)
        
        For i = UBound(arr) To LBound(arr) Step -1
            If arr(i) = Asc("Z") Then
                arr(i) = Asc("A")
                If i = LBound(arr) Then
                    resultado = "A" & StrConv(arr, vbUnicode)
                End If
            Else
                arr(i) = arr(i) + 1
                resultado = StrConv(arr, vbUnicode)
                Exit For
            End If
        Next i
    End If
    
    IncrementarLetras = resultado
End Function
