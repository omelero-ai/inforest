Attribute VB_Name = "DLL3500"
Option Explicit

Public Const OP_VENTA = "01"
Public Const OP_VENTA_CASH = "02"
Public Const OP_VENTA_CUOTAS = "03"
Public Const OP_ANULACION = "04"
Public Const OP_CIERRE = "07"
Public Const OP_SERVICIOS = "08"
Public Const OP_INICIALIZA_TERMINAL = "09"
Public Const OP_SIMULACION_CUOTAS = "10"
Public Const OP_TEST_COMUNICACIONES = "11"
Public Const OP_DUPLICADO = "12"
Public Const OP_DETALLE_OPERACIONES = "13"
Public Const OP_FINANCIERA = "90"
Public Const OP_NO_FINANCIERA = "91"
Public Const RET_RUNNING = 1
Public Const RET_OK = 0
Public Const RET_NOK = -1
Public Const ERR_COM = -2
Public Const ERR_PPAD_NO_RESP = -3
Public Const ERR_TYPE_OPER = -4
Public Const ERR_SOCKET = -5
Public Const ERR_HOST_NO_RESP = -6
Public Const ERR_PRINTER = -7
Public Const ERR_NAK_RECEIVED = -8
Public Const ERR_LRC_PPAD = -9
Public Declare Function fiOpenPort Lib "caja_pinpad.dll" (ConfigFileName As String) As Integer
Public Declare Function fiClosePort Lib "caja_pinpad.dll" () As Integer
Public Declare Function fiStartOperation Lib "caja_pinpad.dll" (pucTipoOperation As String, iTimeOut As Integer, ByRef pucResponse As String) As Integer
Public Declare Function fiGetStatus Lib "caja_pinpad.dll" (ByRef sBuffer As String, ByVal SizeBuffer As Integer) As Integer
Public IpPinPad As String
Public IpPort As Long
Public nTimeOut As Long

Public Const FS = 28

Public Function MensajePinPad(sRetorno As String) As String
    If Mid(sRetorno, 5, 1) = "I" Then
       Printer.Print Mid(sRetorno, 7, Mid(sRetorno, 2, 3) - 2)
       MensajePinPad = "X"
       
    ElseIf Mid(sRetorno, 5, 1) = "C" Then
       Select Case Mid(sRetorno, 7, 1)
              Case "A" To "L"
                   If Mid(sRetorno, 7, 1) = "K" Then
                      Printer.EndDoc
                   Else
                      MensajePinPad = Mid(sRetorno, 7, Mid(sRetorno, 2, 3) - 2)
                   End If
              
              Case Else
                   MsgBox Mid(sRetorno, 7, Mid(sRetorno, 2, 3) - 2), , "Mensaje PinPad"
       End Select
    End If
    
End Function

Public Function ImprimeCabecera(sRetorno As String, Emisor As Boolean) As Boolean
    If Mid(sRetorno, 5, 1) = "I" And Emisor Then
       On Error Resume Next
       Dim Item As String
       Dim sLinea As String
       Dim txt As String
       
       'Cabecera
       ImprimeXCentro sRazonSocial, 40
       ImprimeXCentro sRazonComercial, 40
'       sLinea = String((40 - Len(Trim("R.U.C.: " + sRUC))) / 2, " ") & "R.U.C.: " + sRUC
        Select Case pais
            Case "003" ' argentina
                sLinea = String((40 - Len(Trim("C.U.I.T.: " + sRUC))) / 2, " ") & "C.U.I.T.: " + sRUC
            Case Else ' PERU
                sLinea = String((40 - Len(Trim("R.U.C.: " + sRUC))) / 2, " ") & "R.U.C.: " + sRUC
        End Select
       Printer.Print sLinea
       ImprimeXCentro sDireccion, 40
       sLinea = String((40 - Len(Trim("Telefono: " + sTelefono))) / 2, " ") & "Telefono: " + sTelefono
       Printer.Print sLinea
       Printer.Print " "
       ImprimeCabecera = False
    Else
       ImprimeCabecera = Emisor
    End If
    
End Function

Public Function BuscaRetornoPinPad(sRetorno As String, IdCampo As String) As String
   Dim i As Integer
   Dim j As Integer
   Dim X As Integer
   Dim sClave As String
    
   sClave = Chr$(FS) & IdCampo
   
   If Mid(sRetorno, 1, 1) = IdCampo Then
      i = 0
   Else
      i = InStr(1, sRetorno, sClave, vbTextCompare)
   End If
      
   If i = 0 And Mid(sRetorno, 1, 1) <> IdCampo Then
      BuscaRetornoPinPad = ""
   Else
      X = 0
      For j = i + 2 To Len(sRetorno)
          X = X + 1
          If Mid(sRetorno, j, 1) = Chr$(FS) Then
             Exit For
          End If
      Next j
   End If
   
   If X < 3 Then
      BuscaRetornoPinPad = ""
   Else
      BuscaRetornoPinPad = Mid(sRetorno, i + 2, X - 1)
   End If
End Function
