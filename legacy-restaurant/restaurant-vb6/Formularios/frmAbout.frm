VERSION 5.00
Object = "{0E9D0E41-7AB8-11D1-9400-00A0248F2EF0}#1.0#0"; "dzactx.dll"
Object = "{0FB90DC1-97D1-11D1-87C0-444553540000}#1.0#0"; "duzactx.dll"
Object = "{6C5FD781-9ED8-11D1-87C0-444553540000}#1.0#0"; "dzstactx.dll"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "Acerca del InfoRest"
   ClientHeight    =   7335
   ClientLeft      =   2295
   ClientTop       =   1605
   ClientWidth     =   9720
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0442
   ScaleHeight     =   5062.746
   ScaleMode       =   0  'User
   ScaleWidth      =   9127.583
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Conexiones Actuales"
      Height          =   615
      Left            =   720
      Picture         =   "frmAbout.frx":2415C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   12360
      Width           =   2265
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   435
      Left            =   360
      OleObjectBlob   =   "frmAbout.frx":2459E
      TabIndex        =   5
      Top             =   10800
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton cmdInformacion 
      Caption         =   "Información"
      Height          =   615
      Left            =   3360
      Picture         =   "frmAbout.frx":264C1
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   12240
      Width           =   1425
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   615
      Left            =   5280
      Picture         =   "frmAbout.frx":265C3
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   12360
      Width           =   1425
   End
   Begin VB.Frame Frame 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   90
      Left            =   1440
      TabIndex        =   2
      Top             =   10680
      Width           =   5370
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   300
      Left            =   7920
      TabIndex        =   7
      Top             =   11640
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmAbout.frx":266B5
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmAbout.frx":266D1
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmAbout.frx":266ED
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin VB.Image imgConexiones 
      Height          =   615
      Left            =   8040
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Image imgSalir 
      Height          =   735
      Left            =   8880
      Top             =   120
      Width           =   855
   End
   Begin VB.Image imgInformacion 
      Height          =   615
      Left            =   360
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label lblCaducidad 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fecha de caducidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0073B43C&
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   3840
      Width           =   3735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "20509992461"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0073B43C&
      Height          =   360
      Left            =   360
      TabIndex        =   10
      Top             =   2400
      Width           =   9075
   End
   Begin VB.Image Image3 
      Height          =   1455
      Left            =   1440
      Picture         =   "frmAbout.frx":26709
      Stretch         =   -1  'True
      Top             =   10920
      Width           =   1590
   End
   Begin VB.Label lblFecha 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fecha de compilación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0073B43C&
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   3480
      Width           =   3735
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Versión Ap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0073B43C&
      Height          =   240
      Left            =   3120
      TabIndex        =   8
      Top             =   2956
      Width           =   1155
   End
   Begin DZSTACTXLibCtl.dzstactxctrl dzstactxctrl1 
      Height          =   195
      Left            =   10320
      OleObjectBlob   =   "frmAbout.frx":2D2D7
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin DZACTXLibCtl.dzactxctrl dzactxctrl1 
      Left            =   10680
      OleObjectBlob   =   "frmAbout.frx":2D34F
      Top             =   3240
   End
   Begin DUZACTXLibCtl.duzactxctrl duzactxctrl1 
      Left            =   10680
      OleObjectBlob   =   "frmAbout.frx":2D443
      Top             =   2640
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Razon Social"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0073B43C&
      Height          =   360
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   9075
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAbout.frx":2D53B
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   3120
      TabIndex        =   0
      Top             =   11040
      Width           =   3645
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Conexiones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   13
      Top             =   6960
      Width           =   1095
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub cmdInformacion_Click()
   Call StartSysInfo
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
    frmConexionesActuales.Show vbModal
End Sub

Private Sub Form_Load()
    Centrar Me
    On Error Resume Next
    'lblVersion.Caption = "Versión App: " & App.Major & "." & App.Minor
    'lblRev.Caption = " Rv." & App.Revision
    lblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
    LblFecha.Caption = Format$(FileSystem.FileDateTime(App.path & "\" & App.EXEName & ".exe"), "DD/MM/YYYY")  '"Fecha de compilación:" &
    If sVencimientoLicencia = "00/00/0000" Then
        lblCaducidad.Caption = "Indefinida" ' Fecha de Caducidad:
    Else
         lblCaducidad.Caption = sVencimientoLicencia  ' "Fecha de Caducidad:  " &
    End If
    Label3.Caption = sRazonSocial
    Label5.Caption = sRUC
    
'    If sModulo = "ADICION" Then
'        Me.imgadicion.Visible = True
'    End If
'    If sModulo = "ADMINISTRACION" Then
'        Me.imgAdministracion.Visible = True
'    End If
'    If sModulo = "INFOREST" Then
'        Me.imgPuntoVenta.Visible = True
'    End If
'    If sModulo = "CONSULTA" Then
'        Me.imgConsultas.Visible = True
'    End If
    
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "Información del sistema no disponible...", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frmAbout = Nothing
End Sub


Private Sub imgConexiones_Click()
    frmConexionesActuales.Show vbModal
End Sub

Private Sub imgInformacion_Click()
    Call StartSysInfo
End Sub



Private Sub imgSalir_Click()
    Unload Me
End Sub


