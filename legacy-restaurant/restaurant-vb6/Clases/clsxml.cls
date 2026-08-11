VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsxml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function TraeIp(cSubArea As String) As String
Dim cIp As String

    For i = 3 To Len(cSubArea)
        If Mid(cSubArea, i, 1) <> "\" Then
            cIp = cIp + Mid(cSubArea, i, 1)
        Else
            Exit For
        End If
    Next i
    TraeIp = cIp
End Function
Public Function CrearXMLGuia(ByVal rsReporte As ADODB.Recordset, ByVal Ruta As String, ByVal fcSalida As String, ByValfcArea As String) As Boolean
    Dim of_exist    As Integer
    Dim IsThere     As Integer
    Dim buffer      As OFSTRUCT
    Dim RsXML       As ADODB.Recordset
    Dim NombreXML   As String
    Dim carpeta     As String
    
    Dim sDestino    As String
    Dim lContinua   As Boolean
    Dim sError      As String
   
    Dim a As New Scripting.FileSystemObject
''    Set sVerificaRuta = CreateObject("Scripting.FileSystemObject")
    
    Dim Reply As ICMP_ECHO_REPLY
    Dim lngSuccess As Long
    Dim strIPAddress As String
     
    On Error GoTo Error_CrearXML
    NombreXML = fcSalida & ".xml"
    lContinua = True
    
    Dim dirx As New Scripting.FileSystemObject
'    Set dir = New FileSystemObject
    If dirx.FolderExists(Ruta & "\xml\") = False Then 'para comprobar si existe
        dirx.CreateFolder (Ruta & "\xml\") 'para crear un directorio
    End If
    Set dirx = Nothing

    Ruta = Ruta & "\xml\" & NombreXML
    If a.FileExists(Ruta) = True Then
        If MsgBox("El Archivo " & NombreXML & " ya Existe," & Chr(13) & "¿Desea sobreescribirlo?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
            a.DeleteFile Ruta, True
        Else
            lContinua = False
            Sw = True
        End If
    End If
    If lContinua Then
        CrearXMLGuia = True
        Set RsXML = New ADODB.Recordset
        Set RsXML = Lib.OpenRecordset(" exec Usp_GuiaTransporteXml '" & fcSalida & "'", Cn) 'rsReporte
        If Not RsXML.EOF Then
         RsXML.Save Ruta, adPersistXML
        End If
        Set RsXML = Nothing
       '' MsgBox "Se creo Archivo de Exportación" & Chr(13) & "de datos: " & Ruta, vbInformation, sMensaje
    End If
       Exit Function
Error_CrearXML:
    CrearXMLGuia = False
    MsgBox err.Number & Space(2) & err.Description, vbCritical, sMensaje
End Function
Public Function CrearXML(rsReporte As ADODB.Recordset, Ruta As String) As Boolean
    Dim of_exist    As Integer
    Dim IsThere     As Integer
    Dim buffer      As OFSTRUCT
    Dim NombreXML   As String
    Dim carpeta     As String
    
    Dim sDestino    As String
    Dim lContinua   As Boolean
    Dim sError      As String
    
  '  Dim clsSelXML   As clsSelectXML
    Dim a As New Scripting.FileSystemObject
    
    Dim Reply As ICMP_ECHO_REPLY
    Dim lngSuccess As Long
    Dim strIPAddress As String
     
    On Error GoTo Error_CrearXML
    
    'NombreXML = "RegistroVenta" & fcInicio & fcFinal & ".xml"
    
    lContinua = True
    
    Ruta = Ruta
    
    If a.FileExists(Ruta) = True Then
        If MsgBox("El Archivo " & NombreXML & " ya Existe," & Chr(13) & "¿Desea sobreescribirlo?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
            a.DeleteFile Ruta, True
        Else
            lContinua = False
            Sw = True
        End If
    End If
    
    
    If lContinua Then
        CrearXML = True
  '      Set clsSelXML = New clsSelectXML
        Set RsXML = New ADODB.Recordset
        Set RsXML = rsReporte
        
        ' Set clsSelXML = Nothing
        
        If Not RsXML.EOF Then
            RsXML.Save Ruta, adPersistXML
            
        End If
        Set RsXML = Nothing
    End If
    
   MsgBox "Se creo Archivo de Exportación" & Chr(13) & "de datos: " & Ruta, vbInformation, sMensaje

     Exit Function

Error_CrearXML:
    CrearXML = False
    MsgBox err.Number & Space(2) & err.Description, vbCritical, sMensaje
End Function

Public Function LeeXMLs(codruta As String) As Boolean
'Leer Xml
Dim RsXML       As ADODB.Recordset
Dim nFields     As Integer
Dim prmXml      As ADODB.Parameter
On Error GoTo Error_LeeXMLs
LeeXMLs = True
Set RsXML = New ADODB.Recordset
RsXML.Open codruta, , , , adCmdFile
nFields = 0
While Not RsXML.EOF
    For nFields = 0 To RsXML.Fields.Count - 1
        Set prmXml = New ADODB.Parameter
            prmXml.Name = CStr(RsXML.Fields(nFields).Name)
            prmXml.Type = RsXML.Fields(nFields).Type
            prmXml.value = RsXML.Fields(nFields).value
            prmXml.Size = RsXML.Fields(nFields).DefinedSize
    Next
    Set prmXml = Nothing
    RsXML.MoveNext
Wend
Set frmDocumentoDetalle.grdGrilla.DataSource = RsXML
Set RsXML = Nothing
Exit Function
Error_LeeXMLs:
    Set RsXML = Nothing
    LeeXMLs = False
    MsgBox err.Number & err.Description
''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''
End Function
