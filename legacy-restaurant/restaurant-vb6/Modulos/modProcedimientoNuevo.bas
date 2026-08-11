Attribute VB_Name = "modProcedimientoNuevo"

Private Declare Sub FullQRCode Lib "qrcodelib.dll" (ByVal autoConfigurate As Boolean, ByVal AutoFit As Boolean, ByVal backColor As Long, ByVal barColor As Long, ByVal texto As String, ByVal correctionLevel As TQRCodeECLevel, ByVal encoding As TQRCodeEncoding, ByVal marginpixels As Integer, ByVal moduleWidth As Integer, ByVal Height As Integer, ByVal Width As Integer, ByVal FileName As String)

Private Declare Sub FastQRCode Lib "qrcodelib.dll" (ByVal texto As String, ByVal FileName As String)

Private Declare Function QRCodeLibVer Lib "qrcodelib.dll" ()

'tipo de cambio en impresion
Dim IsqlTC As String
Dim RsParametroTC As Recordset
'-----------------------------
Dim Picture As Image
Public pasa As Boolean

Dim xproducto As String
Dim ximprimeDetallado As Boolean
Dim PropiedadTem As String ' para for de propiedades
Declare Function GetWindowsDirectoryA Lib "kernel32" _
() '   (ByVal lpBuffer As String, ByVal nSize As Long) As Long
   Declare Function GetSystemDirectoryA Lib "kernel32" _
() '   (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Public Sub ImprimeComprobantePagoMesa247(ByVal DocPago As String)
    On Error GoTo fin

fin:
    Exit Sub
End Sub
Public Function ChrBuscaPunto(cadena As String) As Boolean
Dim i As Integer, Pos As Integer
Dim caracter As String
    ChrBuscaPunto = False
    For Pos = 1 To Len(cadena)
        caracter = Mid(cadena, Pos, 1)
        If caracter = "." Then
            ChrBuscaPunto = True
            Exit Function
        End If
    Next
End Function
Public Sub IniciarMotorCashDrow(ByVal Argumentos As String)
On Error GoTo fin:
    'INICIO DE ACTUALIZADOR
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(App.path & "\CashDrow\MotorCashDrow.exe") Then
        'Kill App.Path & "\images\fact.bmp"
        ShellExecute 1, "Open", App.path & "\CashDrow\MotorCashDrow.exe", Argumentos, "", 1

    Else
        MsgBox "La ruta del motor cashdro no se ha encontrado, favor de informar al administrador de sistemas!!!: "
    End If
    
Exit Sub
fin:
MsgBox "CashDro: " & Error
End Sub
