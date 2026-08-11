Attribute VB_Name = "ModPictureBoxCustom"
Public Sub GetTagParts(pbx As PictureBox, ByRef borderHex As String, ByRef textColorHex As String, ByRef textSize As String, ByRef textBold As String, ByRef t1 As String, ByRef t2 As String)
    Dim partes() As String
    If Len(pbx.Tag) > 0 Then
        partes = Split(pbx.Tag, "|")
        If UBound(partes) >= 0 Then borderHex = partes(0) Else borderHex = "BC:#000000"
        If UBound(partes) >= 1 Then textColorHex = partes(1) Else textColorHex = "TC:#000000"
        If UBound(partes) >= 2 Then textSize = partes(2) Else textSize = "TS:8"
        If UBound(partes) >= 3 Then textBold = partes(3) Else textBold = "TB:False"
        If UBound(partes) >= 4 Then t1 = partes(4) Else t1 = ""
        If UBound(partes) >= 5 Then t2 = partes(5) Else t2 = ""
    Else
        borderHex = "BC:#000000"
        textColorHex = "TC:#000000"
        textSize = "TS:8"
        textBold = "TB:False"
        t1 = ""
        t2 = ""
    End If
End Sub
Public Sub PicboxBorder(pbx As PictureBox, hexBorderColor As String)
    Dim borderHex As String, textColorHex As String, textSize As String, textBold As String, t1 As String, t2 As String
    GetTagParts pbx, borderHex, textColorHex, textSize, textBold, t1, t2
    borderHex = "BC:" & hexBorderColor
    pbx.Tag = borderHex & "|" & textColorHex & "|" & textSize & "|" & textBold & "|" & t1 & "|" & t2
    RefreshPictureBox pbx
End Sub
Public Sub PicboxText(pbx As PictureBox, texto1 As String, texto2 As String, Optional hexTextColor As String = "")
    Dim borderHex As String, textColorHex As String, textSize As String, textBold As String, t1 As String, t2 As String
    GetTagParts pbx, borderHex, textColorHex, textSize, textBold, t1, t2
    If hexTextColor <> "" Then
        textColorHex = "TC:" & hexTextColor
    End If
    t1 = texto1
    t2 = texto2
    pbx.Tag = borderHex & "|" & textColorHex & "|" & textSize & "|" & textBold & "|" & t1 & "|" & t2
    RefreshPictureBox pbx
End Sub
Public Sub PicboxTextColor(pbx As PictureBox, hexTextColor As String)
    Dim borderHex As String, textColorHex As String, textSize As String, textBold As String, t1 As String, t2 As String
    GetTagParts pbx, borderHex, textColorHex, textSize, textBold, t1, t2
    textColorHex = "TC:" & hexTextColor
    pbx.Tag = borderHex & "|" & textColorHex & "|" & textSize & "|" & textBold & "|" & t1 & "|" & t2
    RefreshPictureBox pbx
End Sub
Public Sub PicboxTextBold(pbx As PictureBox, isBold As Byte)
    Dim borderHex As String, textColorHex As String, textSize As String, textBold As String, t1 As String, t2 As String
    GetTagParts pbx, borderHex, textColorHex, textSize, textBold, t1, t2
    textBold = "TB:" & CStr(isBold)
    pbx.Tag = borderHex & "|" & textColorHex & "|" & textSize & "|" & textBold & "|" & t1 & "|" & t2
    RefreshPictureBox pbx
End Sub
Public Sub PicboxTextSize(pbx As PictureBox, fontSize As Integer)
    Dim borderHex As String, textColorHex As String, textSize As String, textBold As String, t1 As String, t2 As String
    GetTagParts pbx, borderHex, textColorHex, textSize, textBold, t1, t2
    textSize = "TS:" & CStr(fontSize)
    pbx.Tag = borderHex & "|" & textColorHex & "|" & textSize & "|" & textBold & "|" & t1 & "|" & t2
    RefreshPictureBox pbx
End Sub
Public Sub PicboxBgColor(pbx As PictureBox, hexBackgroundColor As String)
    pbx.AutoRedraw = True
    pbx.ScaleMode = vbPixels
    pbx.backColor = HexToRGB(hexBackgroundColor)
    RefreshPictureBox pbx
End Sub
Public Sub RefreshPictureBox(pbx As PictureBox)
    pbx.AutoRedraw = True
    pbx.ScaleMode = vbPixels
    pbx.Cls   ' Limpia el contenido y repinta con BackColor

    ' Extraer datos desde Tag
    Dim borderHex As String, textColorHex As String, textSize As String, textBold As String, t1 As String, t2 As String
    GetTagParts pbx, borderHex, textColorHex, textSize, textBold, t1, t2

    ' Configurar la fuente según los valores del Tag
    Dim fSize As Integer
    fSize = val(Replace(textSize, "TS:", ""))
    pbx.Font.Size = fSize
    ' Ajusta la negrita: Si la propiedad Bold no es de solo lectura, se puede asignar:
    If UCase(Replace(textBold, "TB:", "")) = "1" Then
        pbx.Font.Bold = True
    Else
        pbx.Font.Bold = False
    End If

    ' Dibujar el texto (si existe) usando el color del texto
    pbx.ForeColor = HexToRGB(Replace(textColorHex, "TC:", ""))
    If (Len(t1) > 0) Or (Len(t2) > 0) Then
        If t2 = "" Then
            pbx.CurrentX = (pbx.ScaleWidth - pbx.TextWidth(t1)) / 2
            pbx.CurrentY = (pbx.ScaleHeight - pbx.TextHeight(t1)) / 2
            pbx.Print t1
        Else
            Dim totalHeight As Single, posY As Single
            totalHeight = pbx.TextHeight(t1) + pbx.TextHeight(t2) + 5
            posY = (pbx.ScaleHeight - totalHeight) / 2

            pbx.CurrentX = (pbx.ScaleWidth - pbx.TextWidth(t1)) / 2
            pbx.CurrentY = posY
            pbx.Print t1

            pbx.CurrentX = (pbx.ScaleWidth - pbx.TextWidth(t2)) / 2
            pbx.CurrentY = posY + pbx.TextHeight(t1) + 5
            pbx.Print t2
        End If
    End If

    ' Dibujar el borde usando el color almacenado en Tag
    Dim bColor As Long
    bColor = HexToRGB(Replace(borderHex, "BC:", ""))
    pbx.ForeColor = bColor
    pbx.Line (0, 0)-(pbx.ScaleWidth - 1, pbx.ScaleHeight - 1), bColor, B
End Sub


Public Function HexToRGB(hexColor As String) As Long
    hexColor = Replace(hexColor, "#", "")
    HexToRGB = RGB(CLng("&H" & Mid(hexColor, 1, 2)), _
                   CLng("&H" & Mid(hexColor, 3, 2)), _
                   CLng("&H" & Mid(hexColor, 5, 2)))
End Function


Public Function PicboxGetText(pbx As PictureBox, Indice As Integer) As String
    Dim borderHex As String, textColorHex As String, textSize As String, textBold As String, t1 As String, t2 As String
    GetTagParts pbx, borderHex, textColorHex, textSize, textBold, t1, t2
    Select Case Indice
        Case 1
            PicboxGetText = t1
        Case 2
            PicboxGetText = t2
        Case Else
            PicboxGetText = ""
    End Select
End Function

Public Function PicboxGetBorderHex(pbx As PictureBox) As String
    Dim borderHex As String, textColorHex As String, textSize As String
    Dim textBold As String, t1 As String, t2 As String

    GetTagParts pbx, borderHex, textColorHex, textSize, textBold, t1, t2

    ' Extraer solo el valor hexadecimal sin el prefijo "BC:"
    PicboxGetBorderHex = Replace(borderHex, "BC:", "")
End Function

