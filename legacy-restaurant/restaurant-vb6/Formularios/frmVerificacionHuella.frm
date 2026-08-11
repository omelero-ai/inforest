VERSION 5.00
Object = "{675D2E3C-5FBF-4F10-8213-EE25D71049AC}#2.3#0"; "sgfplibx.ocx"
Begin VB.Form frmVerificacionHuella 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Huella Dactilar"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpcion 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Salir"
      Height          =   1350
      Index           =   1
      Left            =   3840
      MaskColor       =   &H000000C0&
      Picture         =   "frmVerificacionHuella.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cerrar"
      Top             =   2040
      Width           =   885
   End
   Begin VB.PictureBox HiddenPict 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      Height          =   2775
      Left            =   -120
      Picture         =   "frmVerificacionHuella.frx":00F2
      ScaleHeight     =   2715
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin SGFPLIBXLib.FpLibXCapture FpLibXCapture1 
      Height          =   2655
      Index           =   2
      Left            =   600
      TabIndex        =   4
      Top             =   840
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
   Begin SGFPLIBXLib.FpLibXVerify FpLibXVerify1 
      Left            =   120
      Top             =   720
      _Version        =   131075
      _ExtentX        =   1058
      _ExtentY        =   1058
      _StockProps     =   0
   End
   Begin VB.Image Image1 
      Height          =   2640
      Left            =   720
      Picture         =   "frmVerificacionHuella.frx":1974
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2880
   End
   Begin VB.Label Prompt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmVerificacionHuella"
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

Private Sub cmdLimpiar_Click()
limpiaVariables
End Sub

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
    
   Isql = "select * from vMozo where lActivo = 1 AND tHuella<>'' Order by nBoton"
    If lMCPV And sModulo = "INFOREST" Then
      Isql = "select tCodigoUsuario as Codigo, tResumido, tPassword as tValor, tBandaMagnetica, nTamano=0, nValor=0, thuella from vGrupousuario Where lActivo = 1 And lModulo01 = 1 and tHuella<>'' "
   Else
      Isql = "select * from vMozo where lActivo = 1 and tHuella<>'' Order by nBoton"
   End If
   
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
                            Unload Me
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
MsgBox "Información de " & Mesero & " no Registrado", vbExclamation, sMensaje
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
'    Case 0
'        result = FpLibXCapture1(Index).LiveCapture(Timeout, register_quality)
'        FpLibXCapture1(Index).GetMinutiaeData (g_FirstMinData)
'
'        g_firstStep = True
'        g_secondStep = False
'        g_vefiryResult = False
'
'
'        FpLibXCapture1(1).Clear
'       ' FpLibXCapture1(2).Clear
'        LabelMessage.Caption = ""
'
'    Case 1
'        If g_firstStep = True Then
'            result = FpLibXCapture1(Index).LiveCapture(Timeout, register_quality)
'            FpLibXCapture1(Index).GetMinutiaeData (g_SecondMinData)
'
'            'you can get the score
'            score = FpLibXVerify1.GetMatchingScore(g_FirstMinData, g_SecondMinData)
'
'            result = FpLibXVerify1.Register(g_FirstMinData, g_SecondMinData)
'            If result = True Then
'                g_secondStep = True
'                LabelMessage.Caption = "Patrón de Huellas Correcto. Puede Guardar el Registro"
'                 lGrabaHuella = True
'
'                g_vefiryResult = False
'              '  FpLibXCapture1(2).Clear
'            Else
'                g_secondStep = False
'             '   FpLibXCapture1(2).Clear
'                If FpLibXVerify1.ErrorCode = ERROR_VERIFY_FAKE Then
'                    LabelMessage.Caption = "Retire el dedo y tocar de nuevo , por favor"
'                     lGrabaHuella = False
'
'                Else
'                    LabelMessage.Caption = "Patrón de Huellas Incorrecto, Intente de Nuevo"
'                     lGrabaHuella = False
'
'                End If
'            End If
'
'        End If
'
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
'            If result = True Then
'                g_vefiryResult = True
'                LabelMessage.Caption = "Identical"
'            Else
'                g_vefiryResult = False
'                If FpLibXVerify1.ErrorCode = ERROR_VERIFY_FAKE Then
'                    LabelMessage.Caption = "Release your finger and touch on again, please"
'                Else
'                    'LabelMessage.Caption = "NOT Identical"
'                End If
'
'            End If
        End If
        
    End Select
 
   ' FpLibXCapture1(0).Refresh
   ' FpLibXCapture1(1).Refresh
   FpLibXCapture1(2).Refresh
   Exit Sub
err1:
     Me.FpLibXCapture1(2).Clear
     FpLibXCapture1(2).Refresh
End Sub


