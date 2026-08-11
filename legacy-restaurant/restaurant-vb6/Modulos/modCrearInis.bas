Attribute VB_Name = "modCrearInis"
Public Sub CrearIniTVS()
    On Error GoTo ErrorHandler
        If Len(Trim$(Dir$(App.Path + "\BlueVision.ini"))) Then
            Verificar_Existe = True
        Else
            Verificar_Existe = False
        End If
        
        Do While Not Verificar_Existe
            frmCrearIniTVS.Show vbModal
            If Len(Trim$(Dir$(App.Path + "\BlueVision.ini"))) Then
                Verificar_Existe = True
            Else
                Verificar_Existe = False
            End If
        Loop
    Exit Sub
ErrorHandler:
End Sub

Public Sub CrearIniHardKey()
    On Error GoTo ErrorHandler
        If Len(Trim$(Dir$(App.Path + "\HARDkey.INI"))) Then
            Verificar_Existe = True
        Else
            Verificar_Existe = False
        End If
        
        Do While Not Verificar_Existe
            frmCrearIniHardkey.Show vbModal
            If Len(Trim$(Dir$(App.Path + "\HARDkey.INI"))) Then
                Verificar_Existe = True
            Else
                Verificar_Existe = False
            End If
        Loop
    Exit Sub
ErrorHandler:
End Sub

Public Sub CrearIniInforest()
    On Error GoTo ErrorHandler
        If Len(Trim$(Dir$(App.Path + "\INFOREST.INI"))) Then
            Verificar_Existe = True
        Else
            Verificar_Existe = False
        End If
        
        Do While Not Verificar_Existe
            frmCrearIniInforest.Show vbModal
            If Len(Trim$(Dir$(App.Path + "\INFOREST.INI"))) Then
                Verificar_Existe = True
            Else
                Verificar_Existe = False
            End If
        Loop
    Exit Sub
ErrorHandler:
End Sub

Public Sub CrearIniAlmacen()
    On Error GoTo ErrorHandler
        If Len(Trim$(Dir$(App.Path + "\Almacen.ini"))) Then
            Verificar_Existe = True
        Else
            Verificar_Existe = False
        End If
        
        Do While Not Verificar_Existe
            frmCrearIniAlmacen.Show vbModal
            If Len(Trim$(Dir$(App.Path + "\Almacen.ini"))) Then
                Verificar_Existe = True
            Else
                Verificar_Existe = False
            End If
        Loop
    Exit Sub
ErrorHandler:
End Sub

Public Sub CrearIniInfhotel()
    On Error GoTo ErrorHandler
        If Len(Trim$(Dir$(App.Path + "\Infhotel.ini"))) Then
            Verificar_Existe = True
        Else
            Verificar_Existe = False
        End If
        
        Do While Not Verificar_Existe
            frmCrearIniInfhotel.Show vbModal
            If Len(Trim$(Dir$(App.Path + "\Infhotel.ini"))) Then
                Verificar_Existe = True
            Else
                Verificar_Existe = False
            End If
        Loop
    Exit Sub
ErrorHandler:
End Sub

Public Function VerificarConexionIni(ByVal sRutaDataSource As String, ByVal sMDBCD As String) As Boolean
    On Error GoTo ErrorHandler
    Dim CnPrueba As Connection
    'Configuración
    Set CnPrueba = New Connection
    CnPrueba.Provider = "SQLOLEDB"
    CnPrueba.CursorLocation = adUseServer
    CnPrueba.ConnectionString = "User ID=infhotel" & _
    ";password=4gust1n-fl0r14n" & _
    ";Data Source=" & sRutaDataSource & _
    ";Initial Catalog=" & sMDBCD
    CnPrueba.CommandTimeout = 250
    CnPrueba.Open
    CnPrueba.Close
    VerificarConexionIni = True
    Exit Function
ErrorHandler:
    MsgBox "Conexion a " & sRutaDataSource & ":" & sMDBCD & " no se puedo realizar." & Chr(13) & err.Description, vbExclamation, "Conexion"
    VerificarConexionIni = False
End Function


