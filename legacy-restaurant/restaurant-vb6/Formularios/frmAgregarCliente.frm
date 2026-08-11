VERSION 5.00
Object = "{B1C46850-3E6A-11D2-8FEB-00104B9E07A7}#3.0#0"; "SSDW3AO.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmAgregarCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Cliente"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSFrame SSFrame1 
      Height          =   3480
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   9060
      _Version        =   65536
      _ExtentX        =   15981
      _ExtentY        =   6138
      _StockProps     =   14
      Caption         =   " Datos Principales "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdTexto1 
         Height          =   555
         Left            =   7665
         TabIndex        =   2
         Top             =   405
         Width           =   1275
         _Version        =   196612
         _ExtentX        =   2249
         _ExtentY        =   979
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         CaptionAlignment=   7
         PictureAlignment=   9
      End
      Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdTexto2 
         Height          =   555
         Left            =   7665
         TabIndex        =   3
         Top             =   990
         Width           =   1275
         _Version        =   196612
         _ExtentX        =   2249
         _ExtentY        =   979
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
         CaptionAlignment=   7
         PictureAlignment=   9
      End
      Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdDireccion 
         Height          =   600
         Left            =   7665
         TabIndex        =   4
         Top             =   1575
         Width           =   1275
         _Version        =   196612
         _ExtentX        =   2249
         _ExtentY        =   1058
         _StockProps     =   78
         Caption         =   "Dirección"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmAgregarCliente.frx":0000
         Font3D          =   3
         CaptionAlignment=   4
         PictureAlignment=   0
      End
      Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdTelefono 
         Height          =   555
         Left            =   7665
         TabIndex        =   5
         Top             =   2205
         Width           =   1275
         _Version        =   196612
         _ExtentX        =   2249
         _ExtentY        =   979
         _StockProps     =   78
         Caption         =   "Teléfono"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmAgregarCliente.frx":08DA
         Font3D          =   3
         CaptionAlignment=   4
         PictureAlignment=   0
      End
      Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdZipCode 
         Height          =   555
         Left            =   7665
         TabIndex        =   6
         Top             =   2790
         Width           =   1275
         _Version        =   196612
         _ExtentX        =   2249
         _ExtentY        =   979
         _StockProps     =   78
         Caption         =   "Zip Code"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmAgregarCliente.frx":15B4
         Font3D          =   3
         CaptionAlignment=   4
         PictureAlignment=   1
      End
      Begin VB.Label txtTelefono 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1500
         TabIndex        =   16
         Top             =   2205
         Width           =   2355
      End
      Begin VB.Label txtRC 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1500
         TabIndex        =   15
         Top             =   990
         Width           =   6090
      End
      Begin VB.Label txtDireccion 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1500
         TabIndex        =   14
         Top             =   1575
         Width           =   6090
      End
      Begin VB.Label lbl1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Razón Social :"
         Height          =   195
         Left            =   345
         TabIndex        =   13
         Top             =   405
         Width           =   1035
      End
      Begin VB.Label lbl2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Razón Comercial :"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   990
         Width           =   1290
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dirección :"
         Height          =   195
         Left            =   615
         TabIndex        =   11
         Top             =   1575
         Width           =   765
      End
      Begin VB.Label txtRS 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1500
         TabIndex        =   10
         Top             =   405
         Width           =   6090
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Teléfono :"
         Height          =   195
         Left            =   660
         TabIndex        =   9
         Top             =   2205
         Width           =   720
      End
      Begin VB.Label txtZipCode 
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1500
         TabIndex        =   8
         Top             =   2790
         Width           =   2355
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Zip Code :"
         Height          =   195
         Left            =   645
         TabIndex        =   7
         Top             =   2790
         Width           =   735
      End
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdEmpresa 
      Height          =   555
      Left            =   90
      TabIndex        =   0
      Top             =   3600
      Width           =   1275
      _Version        =   196612
      _ExtentX        =   2249
      _ExtentY        =   979
      _StockProps     =   78
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      CaptionAlignment=   7
      PictureAlignment=   0
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdCancelar 
      Height          =   555
      Left            =   6345
      TabIndex        =   17
      Top             =   3600
      Width           =   1275
      _Version        =   196612
      _ExtentX        =   2249
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Cancela"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmAgregarCliente.frx":228E
      Font3D          =   3
      CaptionAlignment=   1
      PictureAlignment=   9
   End
   Begin SSDataWidgets_A_OLEDB.SSOleDBCommand cmdGrabar 
      Height          =   555
      Left            =   7695
      TabIndex        =   18
      Top             =   3600
      Width           =   1275
      _Version        =   196612
      _ExtentX        =   2249
      _ExtentY        =   979
      _StockProps     =   78
      Caption         =   "Grabar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmAgregarCliente.frx":2B68
      Font3D          =   3
      CaptionAlignment=   1
      PictureAlignment=   9
   End
End
Attribute VB_Name = "frmAgregarCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public wEmpresa As Boolean

Private Sub cmdCancelar_AfterClick()
   Unload Me
End Sub

Private Sub cmdDireccion_AfterClick()
   frmKeyboard.txtResultado.Text = txtDireccion.Caption
   frmKeyboard.Show vbModal
   If wEnter Then
      txtDireccion.Caption = Mid(sDescrip, 1, 50)
   End If
End Sub

Private Sub cmdEmpresa_AfterClick()
   wEmpresa = Not wEmpresa
   CambiaEmpresa (wEmpresa)
End Sub

Private Sub cmdGrabar_AfterClick()
   'Chequea Datos
   If wEmpresa = 1 Then
      If txtRS.Caption = "" Then MsgBox "Ingrese la Razón Social", vbExclamation, sMensaje: Exit Sub
      If txtRC.Caption = "" Then MsgBox "Ingrese la Razón Comercial", vbExclamation, sMensaje: Exit Sub
   Else
      If txtRS.Caption = "" Then MsgBox "Ingrese la Razón Social", vbExclamation, sMensaje: Exit Sub
      If txtRC.Caption = "" Then MsgBox "Ingrese la Razón Comercial", vbExclamation, sMensaje: Exit Sub
   End If
        
   If txtDireccion.Caption = "" Then MsgBox "Ingrese la Dirección", vbExclamation, sMensaje: Exit Sub
   If txtTelefono.Caption = "" Then MsgBox "Ingrese el Teléfono", vbExclamation, sMensaje: Exit Sub
        
   If MsgBox("Seguro de Agregar un nuevo registro?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
      Exit Sub
   End If
   
   'Obtiene el Correlativo
   Set RsCodigo = Lib.OpenRecordset("select CLIENTE from TCODIGO", Cn)
   If IsNull(RsCodigo!cliente) Then
      sCodigo = "00001"
   Else
      sCodigo = Lib.Correlativo(RsCodigo!cliente, 5)
   End If
   
   Cn.Execute "Update TCODIGO Set CLIENTE = '" & sCodigo & "'"
   
   Isql = "insert into TCLIENTE( " & _
          "tCodigoCliente, lEmpresa, tApeCom, tNomSoc, tDireccion, tTelefono1, tZipCode, lActivo, tUsuario, fRegistro) " & _
          "values ('" & sCodigo & "', " & _
                    IIf(wEmpresa, 1, 0) & ", " & _
                  " '" & txtRS.Caption & "', " & _
                  " '" & txtRC.Caption & "', " & _
                  " '" & txtDireccion.Caption & "', " & _
                  " '" & txtTelefono.Caption & "', " & _
                  " '" & txtZipCode.Caption & "', " & _
                                           1 & ", " & _
                  " '" & sUsuario & "', " & _
                  " '" & Date + Time + "')"

           Cn.Execute Isql
           MsgBox "Registro Guardado", vbInformation, sMensaje
           Unload Me
End Sub

Private Sub cmdTelefono_AfterClick()
   frmKeyboard.txtResultado.Text = txtTelefono.Caption
   frmKeyboard.Show vbModal
   If wEnter Then
      txtTelefono.Caption = Mid(sDescrip, 1, 15)
   End If
End Sub

Private Sub cmdTexto1_AfterClick()
   frmKeyboard.txtResultado.Text = txtRS.Caption
   frmKeyboard.Show vbModal
   If wEnter Then
      txtRS.Caption = Mid(sDescrip, 1, 50)
   End If
End Sub

Private Sub cmdTexto2_AfterClick()
   frmKeyboard.txtResultado.Text = txtRC.Caption
   frmKeyboard.Show vbModal
   If wEnter Then
      txtRC.Caption = Mid(sDescrip, 1, 50)
   End If
End Sub

Private Sub cmdZipCode_AfterClick()
   frmKeyboard.txtResultado.Text = txtZipCode.Caption
   frmKeyboard.Show vbModal
   If wEnter Then
      txtZipCode.Caption = Mid(sDescrip, 1, 20)
   End If
End Sub

Private Sub Form_Load()
   Centrar Me
   wEmpresa = True
   CambiaEmpresa (wEmpresa)
End Sub

Public Sub CambiaEmpresa(wflag As Boolean)
   If wflag Then
      cmdEmpresa.Caption = "Cambiar a Cliente"
      lbl1.Caption = "Razón Comercial :"
      lbl2.Caption = "Razón Social :"
      cmdTexto1.Caption = "Razón Comercial"
      cmdTexto2.Caption = "Razón Social"
   Else
      cmdEmpresa.Caption = "Cambiar a Empresa"
      lbl1.Caption = "Apellidos :"
      lbl2.Caption = "Nombres :"
      cmdTexto1.Caption = "Apellidos"
      cmdTexto2.Caption = "Nombres"
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmAgregarCliente = Nothing
End Sub
