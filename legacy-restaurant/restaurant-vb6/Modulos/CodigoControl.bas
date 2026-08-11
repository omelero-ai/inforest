Attribute VB_Name = "CodigoControl"
Private Dihedral(9) As Variant
Private FnF(7) As Variant
Private InverseD5 As Variant

Public Function ObtenerVerhoeff(ByVal Numero As String) As String
Dim tCheck As Integer, i As Integer
'calcula el digito que se aumentara
VerhoeffArrayInit
    For i = Len(Numero) To 1 Step -1
        tCheck = Dihedral(tCheck)(FnF((Len(Numero) - i + 1) Mod 8)(Val(Mid(Numero, i, 1))))
    Next
    ObtenerVerhoeff = IdValue & InverseD5(tCheck)

End Function
Private Sub VerhoeffArrayInit()
'crea las matrices necesarias
Dim i As Integer, j As Integer
'verifica si no esta creada
    If VarType(InverseD5) >= vbArray Then Exit Sub
    'trabajamos con la matris
    Dihedral(0) = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    Dihedral(1) = Array(1, 2, 3, 4, 0, 6, 7, 8, 9, 5)
    Dihedral(2) = Array(2, 3, 4, 0, 1, 7, 8, 9, 5, 6)
    Dihedral(3) = Array(3, 4, 0, 1, 2, 8, 9, 5, 6, 7)
    Dihedral(4) = Array(4, 0, 1, 2, 3, 9, 5, 6, 7, 8)
    Dihedral(5) = Array(5, 9, 8, 7, 6, 0, 4, 3, 2, 1)
    Dihedral(6) = Array(6, 5, 9, 8, 7, 1, 0, 4, 3, 2)
    Dihedral(7) = Array(7, 6, 5, 9, 8, 2, 1, 0, 4, 3)
    Dihedral(8) = Array(8, 7, 6, 5, 9, 3, 2, 1, 0, 4)
    Dihedral(9) = Array(9, 8, 7, 6, 5, 4, 3, 2, 1, 0)
    'crea las matrices
    FnF(0) = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
    FnF(1) = Array(1, 5, 7, 6, 2, 8, 3, 0, 9, 4)
    'calcula la funcion de la matris
    For i = 2 To 7
        FnF(i) = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
        For j = 0 To 9
            FnF(i)(j) = FnF(i - 1)(FnF(1)(j))
        Next
    Next
    'crea el numero inverso
    InverseD5 = Array("0", "4", "3", "2", "1", "5", "6", "7", "8", "9")
End Sub
Public Function allegedrc4(ByVal codigo As String, ByVal llavellegada As String) As String
Dim State(256) As Integer
Dim Mensaje, llave, MsgCif As String
Dim X, Y, Index1, Index2, NMen, i As Integer
    X = 0: Y = 0: Index1 = 0: Index2 = 0
    Mensaje = codigo
    llave = llavellegada
    For i = 0 To 255
        State(i) = i
    Next i
    For i = 0 To 255
        op1 = Asc(Mid(llave, Index1 + 1, 1))
        Index2 = (op1 + State(i) + Index2) Mod 256
        Aux = State(i)
        State(i) = State(Index2)
        State(Index2) = Aux
        Index1 = (Index1 + 1) Mod Len(llave)
    Next i
    For i = 0 To Len(Mensaje) - 1
        X = (X + 1) Mod 256
        Y = (State(X) + Y) Mod 256
        Aux = State(X)
        State(X) = State(Y)
        State(Y) = Aux
        op1 = Asc(Mid(Mensaje, i + 1, 1))
        op2 = State((State(X) + State(Y)) Mod 256)
        NMen = op1 Xor op2
        nrohex = Hex(NMen)
        If Len(nrohex) = 1 Then nrohex = "0" + nrohex
        'MsgCif = MsgCif + "-" + nrohex
        MsgCif = MsgCif + nrohex
    Next i
    'MsgCif = Right(MsgCif, Len(MsgCif) - 1)
    MsgCif = Right(MsgCif, Len(MsgCif))
    allegedrc4 = MsgCif
End Function
Public Function CuantasVeces(ByVal sCadena As String) As Long
    'Esta función devolverá la suma de lo que devuelve el alleged rc4
Dim obtenercaracter As String
Dim Numero As Integer
Dim i As Long
Dim nVeces As Integer
    Numero = 0
    For i = 1 To Len(sCadena)
        obtenercaracter = Mid$(sCadena, i, 1)
        Numero = Asc(obtenercaracter)
        nVeces = Numero + nVeces
    Next
    'Devolver el total de veces
    CuantasVeces = nVeces
End Function
Public Function CuantasVecesde5(ByVal sCadena As String, ByVal inicio As String) As Long
Dim obtenercaracter As String
Dim Numero As Integer
Dim veces As Integer
    Numero = 0
    Do While inicio <= Len(sCadena)
        obtenercaracter = Mid$(sCadena, inicio, inicio)
        Numero = Asc(obtenercaracter)
        veces = Numero + veces
        inicio = inicio + 5
    Loop
    CuantasVecesde5 = veces
End Function
Function ObtenerBase64(ByVal Numero As Long) As String
Dim Diccionario(64) As String
Dim i, j As Integer
Dim Cociente, Resto As Integer
Dim Palabra As String
    j = 0
    For i = 48 To 122
        If i < 58 Then
            j = j + 1
            Diccionario(j) = Chr(i)
        ElseIf i > 64 Then
            If i < 91 Then
                j = j + 1
                Diccionario(j) = Chr(i)
            ElseIf i > 96 Then
                j = j + 1
                Diccionario(j) = Chr(i)
            End If
        End If
     Next i
    Diccionario(63) = "+"
    Diccionario(64) = "/"
    Cociente = 1
    Palabra = ""
    Do While (Cociente > 0)
        Cociente = Numero \ 64
        Resto = Numero Mod 64
        Palabra = Diccionario(Resto + 1) + Palabra
        Numero = Cociente
    Loop
    ObtenerBase64 = Palabra
End Function
Public Function cifrado(ByVal codigo As String, ByVal llavellegada As String) As String
Dim State(256) As Integer
Dim Mensaje, llave, MsgCif As String
Dim X, Y, Index1, Index2, NMen, i As Integer
    X = 0: Y = 0: Index1 = 0: Index2 = 0
    Mensaje = codigo
    llave = llavellegada
    For i = 0 To 255
        State(i) = i
    Next i
    For i = 0 To 255
        op1 = Asc(Mid(llave, Index1 + 1, 1))
        Index2 = (op1 + State(i) + Index2) Mod 256
        Aux = State(i)
        State(i) = State(Index2)
        State(Index2) = Aux
        Index1 = (Index1 + 1) Mod Len(llave)
    Next i
    For i = 0 To Len(Mensaje) - 1
         X = (X + 1) Mod 256
         Y = (State(X) + Y) Mod 256
         Aux = State(X)
         State(X) = State(Y)
         State(Y) = Aux
         op1 = Asc(Mid(Mensaje, i + 1, 1))
         op2 = State((State(X) + State(Y)) Mod 256)
         NMen = op1 Xor op2
         nrohex = Hex(NMen)
         If Len(nrohex) = 1 Then nrohex = "0" + nrohex
         MsgCif = MsgCif + "-" + nrohex
        ' MsgCif = MsgCif + nrohex
    Next i
    MsgCif = Right(MsgCif, Len(MsgCif) - 1)
    'MsgCif = Right(MsgCif, Len(MsgCif))
    cifrado = MsgCif
End Function
Public Function Redondear(dblnToR As Double, Optional intCntDec As Integer) As Double
Dim dblPot As Double
Dim dblF As Double
    If dblnToR < 0 Then
        dblF = -0.5
    Else
        dblF = 0.5
        dblPot = 10 ^ intCntDec
        Redondear = Fix(dblnToR * dblPot * (1 + 1E-16) + dblF) / dblPot
    End If
End Function

Public Function ObtenerCodigoFinal(NAutorizacion, NFactura, NIT, Fecha, MontoTot, Llavedosificacion) As String
Dim codigo1, codigo2, codigo3, codigo4, codigo5, b64 As String
Dim verho1, verho2, verho3, verho4, sumaverhoeff, cadenafinalverhoeff, llavecifradoverhoeff As String
Dim cadenallave1, cadenallave2, cadenallave3, cadenallave4, cadenallave5 As String
Dim unover, dosver, tresver, cuatrover, cincover As String
Dim alleged, resultadoalleged, sumatotalalleged As String
Dim sumaparcial1, sumaparcial2, sumaparcial3, sumaparcial4, sumaparcial5 As String
Dim codigodecontrol As String
    
    
'primer paso
    a = MontoTot
    MontoTot = Redondear(CDbl(a), 0)
    codigo1 = CStr(Val(NFactura) & ObtenerVerhoeff(Val(NFactura)))
    verho1 = CStr(codigo1 & ObtenerVerhoeff(Val(codigo1)))

    codigo2 = CStr(Val(NIT) & ObtenerVerhoeff(Val(NIT)))
    verho2 = CStr(codigo2 & ObtenerVerhoeff(Val(codigo2)))
    If Val(NIT) = 0 Then verho2 = "047"
    
    codigo3 = CStr(Val(Fecha) & ObtenerVerhoeff(Val(Fecha)))
    verho3 = CStr(codigo3 & ObtenerVerhoeff(Val(codigo3)))
    
    codigo4 = CStr(Format(Val(MontoTot), 0) & ObtenerVerhoeff(Format(Val(MontoTot), 0)))
    verho4 = CStr(codigo4 & ObtenerVerhoeff(Val(codigo4)))

    sumaverhoeff = Val(verho1) + Val(verho2) + Val(verho3) + Val(verho4)

    unover = (Val(verho1) + Val(verho2) + Val(verho3) + Val(verho4)) & ObtenerVerhoeff(Val(verho1) + Val(verho2) + Val(verho3) + Val(verho4))
    dosver = Val(unover) & ObtenerVerhoeff(Val(unover))
    tresver = Val(dosver) & ObtenerVerhoeff(Val(dosver))
    cuatrover = Val(tresver) & ObtenerVerhoeff(Val(tresver))
    cincover = CStr(cuatrover) & CStr(ObtenerVerhoeff(cuatrover))


    cadenallave1 = NAutorizacion & Mid((Llavedosificacion), 1, Val(Mid(Right(cincover, 5), 1, 1) + 1))
    cadenallave2 = verho1 & Mid((Llavedosificacion), Val(Mid(Right(cincover, 5), 1, 1) + 1) + 1, Val(Mid(Right(cincover, 5), 2, 1) + 1))
    cadenallave3 = verho2 & Mid((Llavedosificacion), Val(Mid(Right(cincover, 5), 1, 1) + 1) + Val(Mid(Right(cincover, 5), 2, 1) + 1) + 1, Val(Mid(Right(cincover, 5), 3, 1) + 1))
    cadenallave4 = verho3 & Mid((Llavedosificacion), Val(Mid(Right(cincover, 5), 1, 1) + 1) + Val(Mid(Right(cincover, 5), 2, 1) + 1) + Val(Mid(Right(cincover, 5), 3, 1) + 1) + 1 _
      , Val(Mid(Right(cincover, 5), 4, 1) + 1))
    cadenallave5 = verho4 & Mid((Llavedosificacion), Val(Mid(Right(cincover, 5), 1, 1) + 1) + Val(Mid(Right(cincover, 5), 2, 1) + 1) + Val(Mid(Right(cincover, 5), 3, 1) + 1) + Val(Mid(Right(cincover, 5), 4, 1) + 1) + 1 _
      , Val(Mid(Right(cincover, 5), 5, 1) + 1))

    cadenafinalverhoeff = NAutorizacion & Mid((Llavedosificacion), 1, Val(Mid(Right(cincover, 5), 1, 1) + 1)) & _
        verho1 & Mid((Llavedosificacion), Val(Mid(Right(cincover, 5), 1, 1) + 1) + 1, Val(Mid(Right(cincover, 5), 2, 1) + 1)) & _
        verho2 & Mid((Llavedosificacion), Val(Mid(Right(cincover, 5), 1, 1) + 1) + Val(Mid(Right(cincover, 5), 2, 1) + 1) + 1, Val(Mid(Right(cincover, 5), 3, 1) + 1)) & _
        verho3 & Mid((Llavedosificacion), Val(Mid(Right(cincover, 5), 1, 1) + 1) + Val(Mid(Right(cincover, 5), 2, 1) + 1) + Val(Mid(Right(cincover, 5), 3, 1) + 1) + 1 _
                , Val(Mid(Right(cincover, 5), 4, 1) + 1)) & _
        verho4 & Mid((Llavedosificacion), Val(Mid(Right(cincover, 5), 1, 1) + 1) + Val(Mid(Right(cincover, 5), 2, 1) + 1) + Val(Mid(Right(cincover, 5), 3, 1) + 1) + Val(Mid(Right(cincover, 5), 4, 1) + 1) + 1 _
                , Val(Mid(Right(cincover, 5), 5, 1) + 1))
    llavecifradoverhoeff = Llavedosificacion & Right(cincover, 5)
    alleged = allegedrc4(cadenafinalverhoeff, llavecifradoverhoeff)


    sumatotalalleged = CuantasVeces(allegedrc4(cadenafinalverhoeff, llavecifradoverhoeff))
    sumaparcial1 = CuantasVecesde5(allegedrc4(cadenafinalverhoeff, llavecifradoverhoeff), 1)
    sumaparcial2 = CuantasVecesde5(allegedrc4(cadenafinalverhoeff, llavecifradoverhoeff), 2)
    sumaparcial3 = CuantasVecesde5(allegedrc4(cadenafinalverhoeff, llavecifradoverhoeff), 3)
    sumaparcial4 = CuantasVecesde5(allegedrc4(cadenafinalverhoeff, llavecifradoverhoeff), 4)
    sumaparcial5 = CuantasVecesde5(allegedrc4(cadenafinalverhoeff, llavecifradoverhoeff), 5)
    
    sumatotalalleged = Int((CuantasVeces(allegedrc4(cadenafinalverhoeff, llavecifradoverhoeff)) * CuantasVecesde5(allegedrc4(cadenafinalverhoeff, llavecifradoverhoeff), 1)) / (Val(Mid(Right(cincover, 5), 1, 1) + 1))) + _
                            Int((CuantasVeces(allegedrc4(cadenafinalverhoeff, llavecifradoverhoeff)) * CuantasVecesde5(allegedrc4(cadenafinalverhoeff, llavecifradoverhoeff), 2)) / (Val(Mid(Right(cincover, 5), 2, 1) + 1))) + _
                            Int((CuantasVeces(allegedrc4(cadenafinalverhoeff, llavecifradoverhoeff)) * CuantasVecesde5(allegedrc4(cadenafinalverhoeff, llavecifradoverhoeff), 3)) / (Val(Mid(Right(cincover, 5), 3, 1) + 1))) + _
                            Int((CuantasVeces(allegedrc4(cadenafinalverhoeff, llavecifradoverhoeff)) * CuantasVecesde5(allegedrc4(cadenafinalverhoeff, llavecifradoverhoeff), 4)) / (Val(Mid(Right(cincover, 5), 4, 1) + 1))) + _
                            Int((CuantasVeces(allegedrc4(cadenafinalverhoeff, llavecifradoverhoeff)) * CuantasVecesde5(allegedrc4(cadenafinalverhoeff, llavecifradoverhoeff), 5)) / (Val(Mid(Right(cincover, 5), 5, 1) + 1)))
    
    b64 = ObtenerBase64(sumatotalalleged)
    codigodecontrol = cifrado(b64, (Llavedosificacion & Right(cincover, 5)))
    ObtenerCodigoFinal = codigodecontrol
End Function
