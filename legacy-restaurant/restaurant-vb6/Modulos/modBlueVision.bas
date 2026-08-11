Attribute VB_Name = "modBlueVision"
Option Explicit

Private slogin As String
Private sclearPassword As String
Private surl As String
'
Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long
Private Type GUID
   Data1 As Long
   Data2 As Long
   Data3 As Long
   Data4(8) As Byte
End Type

Private Function LeerIniBlueVision()
    slogin = Trim(LeerIni(App.Path + "\BLUEVISION.INI", "BlueVision", "login", "administrador"))
    sclearPassword = Trim(LeerIni(App.Path + "\BLUEVISION.INI", "BlueVision", "ClearPassword", "1111"))
    surl = Trim(LeerIni(App.Path + "\BLUEVISION.INI", "BlueVision", "Url", "http://192.168.3.86:81/"))
End Function

Public Function TVS_EnviarTicket(ByVal RsImpresion As Recordset, sArea As String, ByVal sEvento As String, ByVal sPedido As String) As Boolean
On Error GoTo ErrorHandler
    Call LeerIniBlueVision
    RsImpresion.MoveFirst
    '0- Login
    Dim ycaro As BlueVision_Core_TVS.TvsClientApi
    Dim result As BlueVision_Core_TVS.TvsApiResult
        
    Dim login As BlueVision_Core_TVS.TvsLogin
    Set login = New BlueVision_Core_TVS.TvsLogin
    login.login = slogin
    login.clearPassword = sclearPassword
    Set ycaro = New BlueVision_Core_TVS.TvsClientApi
    ycaro.url = surl
    Set result = ycaro.CreateSession(login)
    
    Dim ticket As BlueVision_Core_TVS.TvsTicket
    Dim linea As BlueVision_Core_TVS.TvsLine
    '1- Creación del ticket
    Set ticket = New BlueVision_Core_TVS.TvsTicket
    ticket.id = Crear_GUID()
    ticket.Mode = "--"
    ticket.Room = sArea
    ticket.Table = RsImpresion!Mesa
    ticket.TicketNumber = 1
    ticket.TOTAL = TVS_ObtenerTotal(RsImpresion, sPedido) 'sTotal
    ticket.TPV = sArea
    ticket.TvsEventName = sEvento
    ticket.Waiter = IIf(IsNull(RsImpresion!Mozo), "SIN MOZO", RsImpresion!Mozo)
    ticket.TimeStamp = Now
    Set result = ycaro.SaveTicket(ticket)
    
    Dim RsProducto As Recordset
    Do While Not RsImpresion.EOF
        Set RsProducto = TVS_ObtenerDescripcionDPedido(RsImpresion, sPedido)
        '2- Creación de linea de ticket 2
        Set linea = New BlueVision_Core_TVS.TvsLine
        linea.Concept = RsImpresion!Producto
        linea.id = Crear_GUID()
        linea.Mode = ticket.Mode '"--"
        linea.Price = RsProducto!nprecioVenta
        linea.Room = ticket.Room 'sArea
        linea.Table = ticket.Table 'RsImpresion!Mesa
        linea.TPV = ticket.TPV 'sArea
        linea.Waiter = ticket.Waiter 'RsImpresion!Mozo
        linea.TicketId = ticket.id
        linea.TOTAL = RsImpresion!nCantidad * RsProducto!nprecioVenta
        linea.Units = RsImpresion!nCantidad
        linea.TimeStamp = Now
        Set result = ycaro.SaveTicketLine(linea)
        RsImpresion.MoveNext
    Loop
    TVS_EscribirLog ("OK " & sArea & " " & sEvento & " " & sPedido)
    Exit Function
ErrorHandler:
    TVS_EscribirLog (err.Description)
End Function

Public Function TVS_EnviarControl(ByVal RsImpresion As Recordset, sArea As String, ByVal sEvento As String, ByVal Texto As String) As Boolean
On Error GoTo ErrorHandler
    Call LeerIniBlueVision
    '0- Login
    Dim ycaro As BlueVision_Core_TVS.TvsClientApi
    Dim result As BlueVision_Core_TVS.TvsApiResult
        
    Dim login As BlueVision_Core_TVS.TvsLogin
    Set login = New BlueVision_Core_TVS.TvsLogin
    login.login = slogin
    login.clearPassword = sclearPassword
    Set ycaro = New BlueVision_Core_TVS.TvsClientApi
    ycaro.url = surl
    Set result = ycaro.CreateSession(login)
    
    Dim Control As New BlueVision_Core_TVS.TvsControl
    Control.id = Crear_GUID()
    Control.TimeStamp = Now
    Control.TPV = sArea
    Control.TvsEventName = sEvento 'this.textBoxControlEventName.Text
    Control.Units = 0 'Int32.Parse(this.textBoxUnitsControl.Text)
    Control.Waiter = RsImpresion!Mozo
    Control.Field1 = Texto
    
    Set result = ycaro.SaveControl(Control)
    If (result.Success = False) Then
        TVS_EscribirLog ("FAIL " & sArea & " " & sEvento & " /" & result.Message)
    Else
        TVS_EscribirLog ("OK " & sArea & " " & sEvento & " /" & result.Message)
    End If
ErrorHandler:
    TVS_EscribirLog (err.Description)
End Function

Private Function TVS_ObtenerDescripcionDPedido(ByVal RsImpresion As Recordset, ByVal sPedido As String) As Recordset
    Isql = "Select * from dPedido Where tCodigoPedido = '" & sPedido & "' And tCodigoProducto IN " & _
           "(Select tCodigoProducto from TProducto Where tDetallado = '" & RsImpresion!Producto & "' OR tResumido = '" & RsImpresion!Producto & "')"
    Set TVS_ObtenerDescripcionDPedido = Lib.OpenRecordset(Isql, Cn)
End Function

Public Function Crear_GUID() As String
    Dim b() As Byte
    Dim lngSize As Long, lngAnswer As Long
    Dim guidData As GUID
    CoCreateGuid guidData
    lngSize = 50
    ReDim b(0 To (lngSize * 2) - 1) As Byte
    lngAnswer = StringFromGUID2(guidData, VarPtr(b(0)), lngSize)
    Crear_GUID = Replace(Replace(Replace(Left(b, lngAnswer - 1), "-", ""), "{", ""), "}", "")
End Function
Private Function TVS_ObtenerTotal(ByVal RsImpresion As Recordset, ByVal sPedido As String) As Double
    Dim sTotal As Double
    sTotal = 0
    RsImpresion.MoveFirst
    Do While Not RsImpresion.EOF
        sTotal = sTotal + RsImpresion!nCantidad * TVS_ObtenerDescripcionDPedido(RsImpresion, sPedido)!nprecioVenta
        RsImpresion.MoveNext
    Loop
    RsImpresion.MoveFirst
    TVS_ObtenerTotal = sTotal
End Function

Private Sub TVS_EscribirLog(ByVal Descripcion As String)
    '*************************************
    'ESCRIBIR ARCHIVOS
    Dim fnum As Integer
    fnum = FreeFile 'se obtiene el numero de archivo
    Open App.Path + "\LogBlueVision.log" For Append As fnum  'se abre para escritura
    
    ' Escribe las lineas de texto
    Print #fnum, Now
    Print #fnum, Descripcion
    Print #fnum, "-------------------"
    
    Close fnum 'se cierra el archivo
End Sub
