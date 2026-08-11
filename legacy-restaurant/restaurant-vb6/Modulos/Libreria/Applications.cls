VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Applications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum Provider
    ACCESS = 1
    SQLSERVER = 2
    EXCEL = 3
    FOXPRO = 4
    DBASE = 5
    VISUAL_FOXPRO = 6
End Enum

Public Function MaxValor(ByVal Tabla As String, ByVal Campo As String, ByVal Conneccion As Connection, Optional Condicion As String) As String
    Dim Rs As New Recordset, Resultado As String
    Rs.CursorLocation = adUseClient
    Isql = "select max(" & Campo & ") as Maximo from " & Tabla & " " & Condicion
    Set Rs = Conneccion.Execute(Isql)
    MaxValor = Rs!Maximo
End Function

Public Function OpenConection(ByVal Proveedor As Provider, ByVal RutaBDD As String, ByVal PassWordBDD As String) As Connection
    Select Case Proveedor
        Case Is = 1
            Set Cn = Nothing
            Set Cn = New Connection
            With Cn
                .CursorLocation = adUseServer
                .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & RutaBDD & ";Persist Security Info=False;Jet OLEDB:Database Password= " & PassWordBDD
                .Open
            End With
        Case Is = 2
            MsgBox "Foxpro" 'Colocar el proveedor de Foxpro para poder levantar tablas de foxpro para DOS
        Case Is = 3
            MsgBox "SqlServer" 'Colocar el Proveedor SQL respectivo
    End Select
    Set OpenConection = Cn
End Function

Public Function OpenRecordset(ByVal Sql As String, ByVal Coneccion As Connection) As Recordset
    Set Rs = New Recordset
    With Rs
        .CursorLocation = adUseClient
        .Open Sql, Coneccion, adOpenDynamic, adLockOptimistic
    End With
    Set OpenRecordset = Rs
End Function

Public Function DiferenciaFechas(ByVal vFec1 As Date, ByVal vFec2 As Date) As String
    ResY = DateDiff("yyyy", vFec1, vFec2)
    ResM = DateDiff("m", vFec1, vFec2)
    ResD = DateDiff("d", vFec1, vFec2)
    
    DiferenciaFechas = IIf(ResY <> "0", IIf(ResY = "1", ResY & " " & "Año", ResY & " " & "Años"), "")
    DiferenciaFechas = IIf(ResM <> "0", IIf(ResM = "1", DiferenciaFechas & "" & ResM & " " & "Mes", DiferenciaFechas & "" & ResM & " " & "Meses"), DiferenciaFechas & "")
    DiferenciaFechas = IIf(ResD <> "0", IIf(ResD = "1", DiferenciaFechas & "" & ResD & " " & "Día", DiferenciaFechas & "" & ResD & " " & "Dias"), DiferenciaFechas & "")
End Function

Public Function NumeroCadena(ByVal strnum As String) As String
    'si se amplia este valor......no se manipularan bien los numeros
    Const cancho = 12
    Const cgrupos = cancho \ 3
    
    If unidad(1) <> "uno" Then
        unidad(1) = "uno"
        unidad(2) = "dos"
        unidad(3) = "tres"
        unidad(4) = "cuatro"
        unidad(5) = "cinco"
        unidad(6) = "seis"
        unidad(7) = "siete"
        unidad(8) = "ocho"
        unidad(9) = "nueve"
        
        decena(1) = "diez"
        decena(2) = "veinte"
        decena(3) = "treinta"
        decena(4) = "cuarenta"
        decena(5) = "cincuenta"
        decena(6) = "sesenta"
        decena(7) = "setenta"
        decena(8) = "ochenta"
        decena(9) = "noventa"
        
        centena(1) = "ciento"
        centena(2) = "doscientos"
        centena(3) = "trescientos"
        centena(4) = "cuatrocientos"
        centena(5) = "quinientos"
        centena(6) = "seiscientos"
        centena(7) = "setecientos"
        centena(8) = "ochocientos"
        centena(9) = "novecientos"
        centena(10) = "cien"
        
        deci(1) = "dieci"
        deci(2) = "veinti"
        deci(3) = "treinta y "
        deci(4) = "cuarenta y "
        deci(5) = "cincuenta y "
        deci(6) = "sesenta y "
        deci(7) = "setenta y "
        deci(8) = "ochenta y "
        deci(9) = "noventa y "
        
        otros(1) = "1"
        otros(2) = "2"
        otros(3) = "3"
        otros(4) = "4"
        otros(5) = "5"
        otros(6) = "6"
        otros(7) = "7"
        otros(8) = "8"
        otros(9) = "9"
        otros(10) = "10"
        otros(11) = "once"
        otros(12) = "doce"
        otros(13) = "trece"
        otros(14) = "catorce"
        otros(15) = "quince"
        
    End If
    
    On Local Error GoTo 0
    
    inga = Abs(strnum)
    'inga = Format(inga, fNumeros)
    negativo = (inga <> CDbl(strnum))
    'strnum = LTrim$(RTrim$(Str$(inga)))
    l = Len(strnum)
    
    If inga < 0 Then
        NumeroCadena = "cero"
        Exit Function
    End If
    
    un = True
    millon = False
    millones = False
    If l < 4 Then un = False
    If inga > 999999 Then millon = True
    If inga > 1999999 Then millones = True
    strb = ""
    strq = strnum
    vez = 0
    
    ReDim strn(1 To cgrupos)
    strq = Right$(String$(cancho, "0") & strnum, cancho)
    For k = Len(strq) To 1 Step -3
        vez = vez + 1
        strn(vez) = Mid$(strq, k - 2, 3)
    Next
    maxvez = cgrupos
    
    For k = cgrupos To 1 Step -1
        If strn(k) = "000" Then
            maxvez = maxvez - 1
        Else
            Exit For
        End If
    Next
    
    For vez = 1 To maxvez
        stru = "": strd = "": strc = ""
        strnum = strn(vez)
        If Left$(strnum, 1) <> "." Then
           l = Len(strnum)
           k = Val(Right$(strnum, 2))
           If Right$(strnum, 1) = "0" Then
               k = k \ 10
               strd = decena(k)
           ElseIf k > 10 And k < 16 Then
               k = Val(Mid$(strnum, l - 1, 2))
               strd = otros(k)
           Else
               stru = unidad(Val(Right$(strnum, 1)))
               If l - 1 > 0 Then
                   k = Val(Mid$(strnum, l - 1, 1))
                   strd = deci(k)
               End If
           End If
           If l - 2 > 0 Then
               k = Val(Mid$(strnum, l - 2, 1))
               If k = 1 Then
                   If Val(strnum) = 100 Then
                       k = 10
                   End If
               End If
               strc = centena(k) & " "
            End If
         Else
            cola = "/100"
            strd = Right$(strnum, 2) & cola
            strd = "y " & strd
            
         End If
         If stru = "" And Left$(strb, 3) = "mil" Then stru = "" Else
              If stru = "uno" And Left$(strb, 9) = "un millon" Then stru = ""
                 strb = strc & strd & stru & " " & strb
         
         'If (vez = 1 Or vez = 3) Then
         If vez = 2 Then
            If strn(vez + 1) <> "000" Then strb = "mil " & strb
         End If
         If vez = 3 And millon Then
            If millones Then
                strb = "millones " & strb
            Else
                strb = "un millón " & strb
            End If
          End If
      Next
      
      strb = Trim$(strb)
      If Right$(strb, 3) = "uno" Then strb = Left$(strb, Len(strb) - 1) & "o"
        'Do   'quitar espacios que hayan de por medio
         '   ia = InStr(strb, " ")
         '   If ia = 0 Then Exit Do
         '       strb = Left$(strb, ia - 1) & Mid$(strb, ia + 1)
        'Loop
        If Left$(strb, 3) = "un " Then strb = Mid$(strb, 5)
        If Left$(strb, 6) = "un mil" Then strb = Mid$(strb, 5)
        If Right$(strb, 16) <> "millones mil una" Then
            ia = InStr(strb, "millones mil una")
            If ia Then strb = Left$(strb, ia + 8) & Mid$(strb, ia + 13)
        End If
        If Right$(strb, 6) = "ciento" Then strb = Left$(strb, Len(strb) - 2)
        If negativo Then strb = "menos" & strb
        
        NumeroCadena = Trim$(strb)
    Exit Function
End Function

Public Function Correlativo(sSecuencia As String, Cant As Integer)
    Dim sSec As String, Cad As String
    Cad = String(Cant, "0")
    sSec = Mid(sSecuencia, 1, Cant)
    sSec = Trim(Str(Val(sSec) + 1))
    Correlativo = Mid(Cad, 1, Cant - Len(sSec)) & sSec
End Function

Public Sub ConfGrid(ByVal CantCol As Integer, ByVal NomGrid As Variant, ParamArray Lista())
    Dim i As Integer
    For i = 0 To CantCol - 1
        If i = 0 Then
           NomGrid.Columns(i).Caption = Lista(i)
           NomGrid.Columns(i).CaptionAlignment = Lista(i + 1)
           NomGrid.Columns(i).DataField = Lista(i + 2)
           NomGrid.Columns(i).Width = Lista(i + 3)
           NomGrid.Columns(i).Alignment = Lista(i + 4)
           NomGrid.Columns(i).Style = Lista(i + 5)
           NomGrid.Columns(i).NumberFormat = Lista(i + 6)
        Else
           NomGrid.Columns(i).Caption = Lista((i * 6) + i)
           NomGrid.Columns(i).CaptionAlignment = Lista((i * 6) + i + 1)
           NomGrid.Columns(i).DataField = Lista((i * 6) + i + 2)
           NomGrid.Columns(i).Width = Lista((i * 6) + i + 3)
           NomGrid.Columns(i).Alignment = Lista((i * 6) + i + 4)
           NomGrid.Columns(i).Style = Lista((i * 6) + i + 5)
           NomGrid.Columns(i).NumberFormat = Lista((i * 6) + i + 6)
        End If
    Next i
End Sub

Public Sub CambioColor(ByVal Color1 As OLE_COLOR, ByVal Color2 As OLE_COLOR, ByVal CantBotones As Integer, ByVal Boton As Integer, ParamArray Botones())
    Dim i As Integer
    For i = 0 To CantBotones - 1
        If i = Boton - 1 Then
            Botones(i).SetFocus
            Botones(i).ForeColor = Color1
        Else
            Botones(i).ForeColor = Color2
        End If
    Next i
End Sub

Public Sub ConfCombo(ByVal CantCol As Integer, ByVal NomCombo As Variant, ParamArray Lista())
    Dim i As Integer
    For i = 0 To CantCol - 1
        If i = 0 Then
           NomCombo.Columns(i).Caption = Lista(i)
           NomCombo.Columns(i).CaptionAlignment = Lista(i + 1)
           NomCombo.Columns(i).DataField = Lista(i + 2)
           NomCombo.Columns(i).Width = Lista(i + 3)
           NomCombo.Columns(i).Alignment = Lista(i + 4)
        Else
           NomCombo.Columns(i).Caption = Lista((i * 4) + i)
           NomCombo.Columns(i).CaptionAlignment = Lista((i * 4) + i + 1)
           NomCombo.Columns(i).DataField = Lista((i * 4) + i + 2)
           NomCombo.Columns(i).Width = Lista((i * 4) + i + 3)
           NomCombo.Columns(i).Alignment = Lista((i * 4) + i + 4)
        End If
    Next i
End Sub



