VERSION 5.00
Object = "{675D2E3C-5FBF-4F10-8213-EE25D71049AC}#2.3#0"; "sgfplibx.ocx"
Begin VB.Form frmVerificarHuellaMarcacion 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Marcación"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox cmdOpcion 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   725
      Index           =   1
      Left            =   1440
      Picture         =   "frmVerificarHuellaMarcacion.frx":0000
      ScaleHeight     =   720
      ScaleWidth      =   2100
      TabIndex        =   9
      Top             =   3840
      Width           =   2100
   End
   Begin VB.Frame FrmRegistroContrasena 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3735
      Left            =   4920
      TabIndex        =   6
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdBusca 
         BackColor       =   &H00F2FAED&
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
         Left            =   4080
         Picture         =   "frmVerificarHuellaMarcacion.frx":1609
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   960
         Width           =   630
      End
      Begin VB.TextBox TxtpasswordUsuario 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   360
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   960
         Width           =   3735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "USER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   360
         TabIndex        =   10
         Top             =   1680
         Width           =   4215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2DEDB&
         Caption         =   "Ingrese su contraseña"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   4575
      End
   End
   Begin VB.Frame FrmHuellaMarcacion 
      BackColor       =   &H00F0F0F0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4815
      Begin VB.PictureBox HiddenPict 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   495
         Left            =   360
         ScaleHeight     =   435
         ScaleWidth      =   555
         TabIndex        =   3
         Top             =   1440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   720
         Picture         =   "frmVerificarHuellaMarcacion.frx":1703
         ScaleHeight     =   255
         ScaleWidth      =   225
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin SGFPLIBXLib.FpLibXCapture FpLibXCapture1 
         Height          =   2895
         Index           =   2
         Left            =   840
         TabIndex        =   4
         Top             =   720
         Width           =   3135
         _Version        =   131075
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   64
         CodeName        =   2
         Brightness      =   28530
         Contrast        =   29287
         Gain            =   28001
      End
      Begin VB.Label Prompt 
         Alignment       =   2  'Center
         BackColor       =   &H00E2DEDB&
         Caption         =   "Presione el Lector de Huella Dactilar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   4575
      End
      Begin VB.Image Image1 
         Height          =   2880
         Left            =   840
         Picture         =   "frmVerificarHuellaMarcacion.frx":BD65D
         Stretch         =   -1  'True
         Top             =   720
         Width           =   3135
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      Height          =   2775
      Left            =   240
      Picture         =   "frmVerificarHuellaMarcacion.frx":CDA9F
      ScaleHeight     =   2715
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   5400
      Visible         =   0   'False
      Width           =   2775
   End
   Begin SGFPLIBXLib.FpLibXVerify FpLibXVerify1 
      Left            =   0
      Top             =   3840
      _Version        =   131075
      _ExtentX        =   1058
      _ExtentY        =   1058
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmVerificarHuellaMarcacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' General error
Const ERROR_NONE = 0
Const ERROR_CREATION_FAILED = 1
Const ERROR_FUNCTION_FAILED = 2
Const ERROR_INVALID_PARAM = 3
Const ERROR_NOT_USED = 4
Const ERROR_DLLLOAD_FAILED = 5
Const ERROR_DLLLOAD_FAILED_DRV = 6
Const ERROR_DLLLOAD_FAILED_ALGO = 7
     
                   
' Device error
Const ERROR_SYSLOAD_FAILED = 51     'system file load fail
Const ERROR_INITIALIZE_FAILED = 52  'chip initialize fail
Const ERROR_LINE_DROPPED = 53       'image data drop
Const ERROR_TIME_OUT = 54           'getliveimage timeout error
Const ERROR_DEVICE_NOT_FOUND = 55   'device not found
Const ERROR_DRVLOAD_FAILED = 56     'dll file load fail
Const ERROR_WRONG_IMAGE = 57         'wrong image
Const ERROR_LACK_OF_BANDWIDTH = 58   'USB Bandwith Lack Error
Const ERROR_DEV_ALREADY_OPEN = 59     'Device Exclusive access Error
Const ERROR_GETSN_FAILED = 60         'Fail to get Device Serial Number
Const ERROR_UNSUPPORTED_DEV = 61      ' Unsupported device
   
   
' Extract&verification error
Const ERROR_FEAT_NUMBER = 101               'too small number of minutiae
Const ERROR_INVALID_TEMPLATE_TYPE = 102     'wrong template type
Const ERROR_INVALID_TEMPLATE1 = 103         'error in decoding template 1
Const ERROR_INVALID_TEMPLATE2 = 104         'error in decoding template 2
Const ERROR_EXTRACT_FAIL = 105
Const ERROR_MATCH_FAIL = 106

' Template Format
Const TEMPLATE_FORMAT_ANSI378 = 256
Const TEMPLATE_FORMAT_SG400 = 512
Const TEMPLATE_FORMAT_ISO19794 = 768


Dim g_FirstMinData() As Byte
Dim g_SecondMinData() As Byte
Dim g_MatchingMinData() As Byte
Dim g_firstStep As Boolean
Dim g_secondStep As Boolean
Dim g_vefiryResult As Boolean

Public WithEvents Capture As DPFPCapture
Attribute Capture.VB_VarHelpID = -1
Dim CreateFtrs As DPFPFeatureExtraction
Dim Verify As DPFPVerification
Dim ConvertSample As DPFPSampleConversion
Dim Templ As DPFPTemplate

Dim RsMozo As New ADODB.Recordset
Dim RsMarca As ADODB.Recordset

Private Sub cmdBusca_Click()
    frmKeyBoard.txtResultado = ""
    frmKeyBoard.Show vbModal
    TxtpasswordUsuario.Text = IIf(wEnter, sDescrip, TxtpasswordUsuario.Text)
    
    If wEnter Then
        If ValidaPassWordUsuarios(Trim(sDescrip)) = False Then
            MsgBox "Usuario No Encontrado", vbExclamation, sMensaje
            TxtpasswordUsuario.Text = ""
        End If
    End If
    
End Sub
Private Function ValidaPassWordUsuarios(ByVal Pass As String) As Boolean
    RsMozo.MoveFirst
    If Not (RsMozo.EOF Or RsMozo.BOF) Then
        ValidaPassWordUsuarios = False
        For X = 1 To RsMozo.RecordCount
                If Desencapsula(RsMozo!tvalor) = Pass Then
                    sVar1 = RsMozo!tResumido
                    sMozo = RsMozo!codigo
                    lSomelier = IIf(IsNull(RsMozo!nValor), 0, RsMozo!nValor)
                      wEnterHuella = True
                      Dim Result As String
                      Set RsMarca = Lib.OpenRecordset(" USP_ADD_MARCACION '" & RsMozo!tvalor & "'", Cn)
                       
                      If Not (RsMarca.EOF Or RsMarca.BOF) Then
                            If RsMarca.Fields(0) = "Marcación Registrada Correctamente" Then
                                frmMarca.resultado = RsMarca.Fields(0)
                                frmMarca.descripcion = RsMozo!descripcion
                              'frmMarca.backColor = &H808000
                              frmMarca.Show vbModal
                               'MsgBox (RsMarca.Fields(0))
                               ValidaPassWordUsuarios = True
                               Unload Me
                               Exit Function
                            Else
                              frmMarca.resultado = RsMarca.Fields(0)
                              frmMarca.descripcion = "Error en registro, volver a ingresar password!!!"
                              'frmMarca.backColor = &H8080FF
                              frmMarca.Show vbModal
                              ValidaPassWordUsuarios = False
                              Exit Function
                            End If
                    Else
                        ValidaPassWordUsuarios = False
                        Exit Function
                    End If
                End If
            RsMozo.MoveNext
        Next X
    End If
End Function



Private Sub cmdOpcion_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Initialize()
   
    Set RsMozo = New ADODB.Recordset
End Sub
Private Sub TemplateFormat_Click(Index As Integer)

    g_firstStep = False
    g_secondStep = False
    g_vefiryResult = False

    ' Set format of template to use.
    ' two modes of FpLibXVerify1 and FpLibXCapture1 MUST be the same minutiae mode
    ' otherwise comment out to use MinutiaeMode property
    Select Case Index
    
        
    Case 1
        FpLibXVerify1.MinutiaeMode = TEMPLATE_FORMAT_ANSI378        ' ANSI 378
'        FpLibXCapture1(0).MinutiaeMode = TEMPLATE_FORMAT_ANSI378
'        FpLibXCapture1(1).MinutiaeMode = TEMPLATE_FORMAT_ANSI378
        FpLibXCapture1(2).MinutiaeMode = TEMPLATE_FORMAT_ANSI378
        
  
    End Select

    'Redim these variables after setting the MinutiaeMode,
    'because size of data is different for each MinutiaeMode
    ReDim g_FirstMinData(FpLibXCapture1(0).MinutiaeSize) As Byte
    ReDim g_SecondMinData(FpLibXCapture1(1).MinutiaeSize) As Byte
    ReDim g_MatchingMinData(FpLibXCapture1(2).MinutiaeSize) As Byte


End Sub
Private Sub Form_Load()
    
    If lPasswordMarcacion Then
        FrmHuellaMarcacion.Visible = False
        FrmRegistroContrasena.Top = 0
        FrmRegistroContrasena.Left = 0
    Else
        FrmRegistroContrasena.Visible = False
    End If
    'FrmRegistroContrasena
    

    If lHuellaDigitalPersona = True Then
        Image1.Visible = True
        FpLibXCapture1(2).Visible = False
        Prompt.Caption = "Presione el Lector de Huella Dactilar"
    End If
    If lHuellaSecugen = True Then
        Prompt.Caption = "Click en la Cuadro Central y Presione el Lector de Huella Dactilar"
        Image1.Visible = False
        FpLibXCapture1(2).Visible = True
        g_firstStep = False
        g_secondStep = False
        g_vefiryResult = False
        FpLibXCapture1(2).CodeName = 2
        g_secondStep = True
        ' hardware initialize
        ' FDP02(0), FDU02(1), FDU03(2), FDU04(3)
        Dim DefaultDevice As Integer
        DefaultDevice = 2   ' FDU03

        FpLibXVerify1.MinutiaeMode = TEMPLATE_FORMAT_ANSI378        ' ANSI 378
        FpLibXCapture1(2).MinutiaeMode = TEMPLATE_FORMAT_ANSI378
        ReDim g_MatchingMinData(FpLibXCapture1(2).MinutiaeSize) As Byte
        FpLibXCapture1(2).Refresh
    End If
    
   Isql = "select codigo, Descripcion, tResumido , nBoton , isnull(tvalor,'') tvalor, lActivo, nValor, tBandaMagnetica, nTamano , isnull(tHuella,'') tHuella , lAutorizaEnviosMozo from vMozo where lActivo = 1 " & _
   "union all " & _
   " select tCodigoUsuario ,tDetallado , tResumido,''nBoton, isnull(tPassword,'') tvalor  , lActivo , ''nvalor, tBandaMagnetica , '' Ntamano, isnull(tHuella,'') tHuella , '' lAutorizaEnviosMozo  from TUSUARIO where lActivo = 1 "
   
   Set RsMozo = Lib.OpenRecordset(Isql, Cn)
   
    If lHuellaDigitalPersona = True Then
                If Capture Is Nothing Then
                    Set Capture = New DPFPCapture
                    Capture.StartCapture
                Else
                    Capture.StartCapture
                End If
                
                Set CreateFtrs = New DPFPFeatureExtraction
                Set Verify = New DPFPVerification
                Set ConvertSample = New DPFPSampleConversion
     End If
End Sub

Public Sub hextoarray(inphex As String, outarray() As Byte)

ReDim outarray(0 To Len(inphex) / 2)

Dim i As Integer
    For i = 1 To Len(inphex) Step 2
    outarray(((i + 1) / 2) - 1) = val("&H" + Mid$(inphex, i, 2))
    Next i
End Sub
Private Sub DrawPicture(ByVal Pict As IPictureDisp)
 ' Must use hidden PictureBox to easily resize picture.
 Set HiddenPict.Picture = Pict
 Picture1.PaintPicture HiddenPict.Picture, _
       0, 0, Picture1.ScaleWidth, _
       Picture1.ScaleHeight, _
       0, 0, HiddenPict.ScaleWidth, _
       HiddenPict.ScaleHeight, vbSrcCopy
 Picture1.Picture = Picture1.Image
End Sub

Private Sub Capture_OnComplete(ByVal ReaderSerNum As String, ByVal Sample As Object)
 Dim Feedback As DPFPCaptureFeedbackEnum
 Dim Res As DPFPVerificationResult
 RsMozo.Requery
 Dim cadena As String
 Dim X As Integer
  Dim blob() As Byte
 DrawPicture ConvertSample.ConvertToPicture(Sample)
 Sleep (500)
 Feedback = CreateFtrs.CreateFeatureSet(Sample, DataPurposeVerification)

 
 If Feedback = CaptureFeedbackGood Then
    If Not (RsMozo.EOF Or RsMozo.BOF) Then
        RsMozo.MoveFirst
        For X = 1 To RsMozo.RecordCount
                cadena = RsMozo!thuella
                hextoarray cadena, blob()
                If Templ Is Nothing Then Set Templ = New DPFPTemplate
                Templ.Deserialize blob
                Set Res = Verify.Verify(CreateFtrs.FeatureSet, Templ)
                
                If Res.Verified = True Then
                    sVar1 = RsMozo!tResumido
                    sMozo = RsMozo!codigo
                    lSomelier = IIf(IsNull(RsMozo!nValor), 0, RsMozo!nValor)
                      
                      wEnterHuella = True
                      
                      Dim Result As String
                  
                      Set RsMarca = Lib.OpenRecordset(" USP_ADD_MARCACION '" & RsMozo!tvalor & "'", Cn)
                       
                      If Not (RsMarca.EOF Or RsMarca.BOF) Then
                       
                            If RsMarca.Fields(0) = "Marcación Registrada Correctamente" Then
                              frmMarca.lblRestultado.Caption = RsMarca.Fields(0)
                              frmMarca.lblNom.Caption = RsMozo!descripcion
                              frmMarca.backColor = &H808000
                              frmMarca.Show vbModal
                               'MsgBox (RsMarca.Fields(0))
                               Unload Me
                            Else
                              frmMarca.lblRestultado.Caption = RsMarca.Fields(0)
                              frmMarca.lblNom.Caption = ""
                              frmMarca.backColor = &H8080FF
                              frmMarca.Show vbModal
                               Exit Sub
                            End If
                       End If
                End If
            RsMozo.MoveNext
        Next X
    End If
  End If
  
  If wEnterHuella = False Then
        limpiaVariables
   End If
 
End Sub


Public Sub limpiaVariables()
Picture1.Picture = Nothing

Capture.StartCapture
Me.FpLibXCapture1(2).Clear
sMozo = ""
lSomelier = False
'MsgBox "Información de Mesero y/o Personal no Registrado", vbExclamation, sMensaje
 frmMarca.lblRestultado.Caption = "No Existe Usuario, favor de validar"
frmMarca.lblNom.Caption = ""
frmMarca.backColor = &H8080FF
frmMarca.Show vbModal


Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If lHuellaDigitalPersona = True Then
        Capture.StopCapture
    End If
    
End Sub


Private Sub FpLibXCapture1_Paint(Index As Integer, ByVal hwnd As stdole.OLE_HANDLE, ByVal hdc As stdole.OLE_HANDLE)
    Select Case Index
   
    Case 2
        If g_firstStep = True And g_secondStep = True And g_vefiryResult = False Then
            TextOut hdc, 30, 70, "Click Aquí..", 12
        End If
    
    End Select
End Sub
Private Sub FpLibXCapture1_Click(Index As Integer)
On Error GoTo err1
Dim register_quality, verify_quality As Integer
Dim Timeout As Integer
Dim j As Integer
Dim cadena1 As String
Dim huella() As Byte
register_quality = 50
verify_quality = 30
Timeout = 6000 ' 6seconds

    Select Case (Index)
    Case 2
        If g_secondStep = True Then
            RsMozo.Requery
            bResult = FpLibXCapture1(Index).LiveCapture(Timeout, verify_quality)
            FpLibXCapture1(Index).GetMinutiaeData (g_MatchingMinData)
                        
            If Not (RsMozo.EOF Or RsMozo.BOF) Then
                    RsMozo.MoveFirst
                    For j = 1 To RsMozo.RecordCount
                        cadena1 = RsMozo!thuella
                        hextoarray cadena1, huella()
                        Result = FpLibXVerify1.Verify(huella, g_MatchingMinData)
                         
                        If Result = True Then
                          sVar1 = RsMozo!tResumido
                            sMozo = RsMozo!codigo
                          lSomelier = IIf(IsNull(RsMozo!nValor), 0, RsMozo!nValor)
                            
                            wEnterHuella = True
                            Unload Me
                        End If
                        RsMozo.MoveNext
                    Next j
           End If
            
             If Result = False Then
            Me.FpLibXCapture1(2).Clear
            End If
        End If
        
    End Select
   FpLibXCapture1(2).Refresh
   Exit Sub
err1:
     Me.FpLibXCapture1(2).Clear
     FpLibXCapture1(2).Refresh
End Sub

Private Sub TxtpasswordUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If ValidaPassWordUsuarios(Trim(TxtpasswordUsuario.Text)) = False Then
            MsgBox "Usuario No Encontrado", vbExclamation, sMensaje
            TxtpasswordUsuario.Text = ""
        End If
   End If
End Sub
