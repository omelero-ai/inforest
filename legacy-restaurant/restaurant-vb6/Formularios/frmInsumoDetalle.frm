VERSION 5.00
Begin VB.Form frmInsumoDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2805
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   11070
   Icon            =   "frmInsumoDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   11070
   Begin VB.Frame fraDetalle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   1740
      TabIndex        =   16
      Top             =   0
      Width           =   9315
      Begin VB.CommandButton cmdBuscaTipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
         Picture         =   "frmInsumoDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   960
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.OptionButton optProducto 
         Caption         =   "Plato de Venta"
         Height          =   495
         Left            =   4920
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optInsumo 
         Caption         =   "Insumo"
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtStock 
         Height          =   285
         Left            =   7800
         MaxLength       =   8
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtMensaje 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   960
         Width           =   4770
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   390
         Width           =   1290
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   405
         TabIndex        =   3
         Top             =   1560
         Width           =   960
      End
      Begin VB.Label lblstock 
         Caption         =   "Disponible Al Momento:"
         Height          =   495
         Left            =   6720
         TabIndex        =   20
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Insumo:"
         Height          =   195
         Left            =   405
         TabIndex        =   18
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   375
         TabIndex        =   17
         Top             =   435
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   11010
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2055
      Width           =   11070
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   8640
         Picture         =   "frmInsumoDetalle.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   9840
         Picture         =   "frmInsumoDetalle.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   6270
         Picture         =   "frmInsumoDetalle.frx":0738
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   7440
         Picture         =   "frmInsumoDetalle.frx":0C6A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   6150
         TabIndex        =   15
         Top             =   60
         Width           =   6210
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmInsumoDetalle.frx":119C
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmInsumoDetalle.frx":16DE
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmInsumoDetalle.frx":1C20
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5670
            Picture         =   "frmInsumoDetalle.frx":2162
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   5160
            Picture         =   "frmInsumoDetalle.frx":26A4
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4680
            Picture         =   "frmInsumoDetalle.frx":2BE6
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            Caption         =   "Registro 0 de 0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1470
            TabIndex        =   19
            Top             =   180
            Width           =   3105
         End
      End
   End
   Begin VB.Image Image 
      Height          =   2025
      Left            =   0
      Picture         =   "frmInsumoDetalle.frx":3128
      Stretch         =   -1  'True
      Top             =   30
      Width           =   1710
   End
End
Attribute VB_Name = "frmInsumoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CGMiranda-------------------------------------------------------------------
Dim iStock As Double
Option Explicit

Sub Asignar()
    iStock = 0
    With frmInsumo.RsCabecera
        'Cuadro de Texto
        Me.optInsumo.value = IIf(IsNull(!LINSUMO), 1, IIf(!LINSUMO = True, 1, 0))
        If optInsumo.value = False Then: Me.optProducto.value = True
        txtCodigo = IIf(IsNull(!tCodigo), "", !tCodigo)
        txtMensaje = IIf(IsNull(!descripcion), "", !descripcion)
        If sModulo = "INFOREST" Then
            txtStock.Text = IIf(IsNull(!nStock), "", !nStock)
            'iStock = Val(txtStock.Text)
        End If
        'Check Box
        chkActivo = IIf(IsNull(!lActivo), 0, IIf(!lActivo = True, 1, 0))
 
    End With
    
End Sub



Private Sub cmdBuscaTipo_Click()
              Me.txtMensaje.Text = ""
               Isql = "select Codigo, Descripcion, Grupo, SubGrupo from vProducto where lActivo=1  order by Descripcion"
               frmBusca.nPredeterm = 3
               Call ConfGrilla(4, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 800, 2, 0, "", _
                                                      "Grupo", 2, "Grupo", 1300, 0, 0, "", _
                                                      "SubGrupo", 2, "SubGrupo", 1400, 0, 0, "", _
                                                      "Plato", 2, "Descripcion", 3900, 0, 0, "")
               frmBusca.Show vbModal
               If Not wEnter Then
                  Exit Sub
               End If
              Me.txtMensaje.Text = sDescrip
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmInsumo.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmInsumo.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmInsumo.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmInsumo.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmInsumo.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmInsumo.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmInsumo.RsCabecera.AbsolutePosition & " de " & frmInsumo.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    'Dim nCorrela As String
    Dim oComando As clsComando

   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               chkActivo.value = 1
               'Cambia el Nombre del Primer Text
             Me.optInsumo.SetFocus
                    
          Case Is = 1 ' Grabar

                'frmInsumo.RsCabecera
               Dim nCorrela As String
               Dim nPos As Integer
                              Dim cont As Integer
                Set oComando = New clsComando
               'Chequea Datos
               'If txtMensaje.Text = "" Then MsgBox " Ingrese el Insumo ", vbExclamation, sMensaje: txtMensaje.SetFocus: Exit Sub
               If txtMensaje.Text = "" Then
                   MsgBox "Ingrese el Insumo", vbExclamation, sMensaje
                   If Me.txtMensaje.Enabled = True Then
                        txtMensaje.SetFocus
                   End If
                   Exit Sub
               End If
                                  
               If Sw Then

                 cont = Calcular("select count(*) as codigo from TINSUMO where Descripcion = UPPER('" & txtMensaje.Text & "')", Cn)
                 If (CInt(cont) > 0) Then
                    MsgBox "Ya existe un insumo/plato con el nombre:  " & txtMensaje.Text & "", vbExclamation, sMensaje
                    Exit Sub
                 End If
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tcodigo) as Codigo from TINSUMO ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      nCorrela = "00000001"
                      txtCodigo.Text = nCorrela
                  Else
                       nCorrela = Lib.Correlativo(nCorrela, 8)
                       txtCodigo.Text = nCorrela
                  End If
                  Sw = False
                txtMensaje.Text = UCase(txtMensaje.Text)
                
                
                'Llama al procedimiento para agregar un mensaje
                Isql = "usp_agregarinsumos"
                  If Not oComando.CreateCmdSp(Isql, Cn) Then
                        Set oComando = Nothing
                        Exit Sub
                    End If
                
                oComando.CreateParameter "@codigo", adVarChar, adParamInput, 8, txtCodigo.Text
                oComando.CreateParameter "@usuario", adVarChar, adParamInput, 15, sUsuario
                oComando.CreateParameter "@descripcion", adVarChar, adParamInput, 50, txtMensaje.Text
                oComando.CreateParameter "@nstock", adDouble, adParamInput, 3, IIf(IsNull(txtStock.Text), 0, val(txtStock.Text))
                oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 3, sCaja
                oComando.CreateParameter "@activo", adBoolean, adParamInput, 1, chkActivo.value
                oComando.CreateParameter "@LINSUMO", adBoolean, adParamInput, 1, Me.optInsumo.value
                  If Not oComando.GetParamOK Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  If Not oComando.ExecSP Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  frmInsumo.RsCabecera.Filter = ""
                  frmInsumo.RsCabecera.Sort = "Descripcion ASC"
                  frmInsumo.RsCabecera.Requery
                  frmInsumo.RsCabecera.MoveLast
                  ActivarBotones (True)
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  cmdTexto.Caption = "Registro " & IIf(frmInsumo.RsCabecera.RecordCount = 0, 0, frmInsumo.RsCabecera.AbsolutePosition) & " de " & frmInsumo.RsCabecera.RecordCount

               Else
                 cont = Calcular("select count(*) as codigo from TINSUMO where Descripcion = UPPER('" & txtMensaje.Text & "') AND tcodigo <> (select tcodigo from TINSUMO where tcodigo = UPPER('" & txtCodigo.Text & "') )", Cn)
                 If (CInt(cont) > 0) Then
                    MsgBox "Ya existe un insumo/plato con el nombre:  " & txtMensaje.Text & "", vbExclamation, sMensaje
                    Exit Sub
                 End If
                    
                    iStock = val(txtStock.Text)
                    txtMensaje.Text = UCase(txtMensaje.Text)
                    Isql = "USP_MODIFICARINSUMOS"
                    If Not oComando.CreateCmdSp(Isql, Cn) Then
                        Set oComando = Nothing
                        Exit Sub
                    End If
                oComando.CreateParameter "@codigo", adVarChar, adParamInput, 8, txtCodigo.Text
                oComando.CreateParameter "@usuario", adVarChar, adParamInput, 15, sUsuario
                oComando.CreateParameter "@descripcion", adVarChar, adParamInput, 50, txtMensaje.Text
                oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 3, sCaja
                oComando.CreateParameter "@activo", adBoolean, adParamInput, 1, chkActivo.value
                oComando.CreateParameter "@nstock", adDouble, adParamInput, 1, IIf(IsNull(txtStock.Text), 0, val(txtStock.Text))
    

                Select Case sModulo
                    Case "INFOREST"
                            oComando.CreateParameter "@MODULO", adBoolean, adParamInput, 1, 1
                    Case "ADMINISTRACION"
                            oComando.CreateParameter "@MODULO", adBoolean, adParamInput, 1, 0
                    Case Else
                End Select
                oComando.CreateParameter "@LINSUMO", adBoolean, adParamInput, 1, Me.optInsumo.value
                    If Not oComando.GetParamOK Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  If Not oComando.ExecSP Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                    
                   Cn.Execute "update tinsumo set lNuevoModificado=1 where tCodigo='" & txtCodigo.Text & "'"
                   
                   nPos = frmInsumo.RsCabecera.AbsolutePosition
                   frmInsumo.RsCabecera.Filter = ""
                   frmInsumo.RsCabecera.Requery
                   frmInsumo.RsCabecera.AbsolutePosition = nPos
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          Case Is = 2 '
                    Set oComando = New clsComando
                    
                If frmInsumo.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
                'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Insumo?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
                    Isql = "USP_ELIMINARINSUMOS"
                    If Not oComando.CreateCmdSp(Isql, Cn) Then
                        Set oComando = Nothing
                        Exit Sub
                    End If
                    oComando.CreateParameter "@codigo", adVarChar, adParamInput, 8, txtCodigo.Text
                   
                    If Not oComando.GetParamOK Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  If Not oComando.ExecSP Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                    frmInsumo.RsCabecera.Requery
                    If frmInsumo.RsCabecera.RecordCount <> 0 Then
                            frmInsumo.RsCabecera.MoveLast
                            Asignar
                            cmdTexto.Caption = "Registro " & IIf(frmInsumo.RsCabecera.RecordCount = 0, 0, frmInsumo.RsCabecera.AbsolutePosition) & " de " & frmInsumo.RsCabecera.RecordCount
                    Else
                            ActivarBotones False
                            Blanquear Me
                            Sw = True
                    End If
          Case Is = 3 ' Salir
               Unload Me
   End Select

End Sub
Public Sub ValidaStr(Cad As TextBox)
  Dim sCad As String
  Dim nRes As Integer
  nRes = 0
  sCad = Trim(Cad.Text)
  nRes = InStr(1, sCad, "&")
  If nRes > 0 Then
     MsgBox "Caracter (&) no permitido, favor corregir", vbExclamation, sMensaje
     Cad.SelStart = nRes - 1
     Cad.SetFocus
     Exit Sub
  End If
  
  nRes = InStr(1, sCad, "'")
  If nRes > 0 Then
     MsgBox "Caracter (') no permitido, favor corregir", vbExclamation, sMensaje
     Cad.SelStart = nRes - 1
     Cad.SetFocus
     Exit Sub
  End If
  
  nRes = InStr(1, sCad, """")
  If nRes > 0 Then
     MsgBox "Caracter ("") no permitido, favor corregir", vbExclamation, sMensaje
     Cad.SelStart = nRes - 1
     Cad.SetFocus
     Exit Sub
  End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Insumos/Platos de Stock Crítico"
    fraDetalle.Caption = Me.Caption
    lblstock.Visible = False
    txtStock.Visible = False
    If sModulo = "INFOREST" Then
        lblstock.Visible = True
        txtStock.Visible = True
    End If
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.value = 1
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & frmInsumo.RsCabecera.AbsolutePosition & " de " & frmInsumo.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmInsumoDetalle = Nothing
End Sub

Sub ActivarBotones(ByVal Activa As Boolean)
    cmdNavegar(0).Enabled = Activa
    cmdNavegar(1).Enabled = Activa
    cmdNavegar(2).Enabled = Activa
    cmdNavegar(3).Enabled = Activa
    cmdNavegar(4).Enabled = Activa
    cmdNavegar(5).Enabled = Activa
    cmdOpcion(0).Enabled = Activa
    cmdOpcion(2).Enabled = Activa
End Sub

Private Sub optInsumo_Click()
    Me.txtStock.MaxLength = 8
    Me.txtStock.Text = ""
    Me.txtMensaje.Text = ""
    Me.txtMensaje.Enabled = True
    cmdBuscaTipo.Visible = False
End Sub

Private Sub optProducto_Click()
    Me.txtStock.MaxLength = 8
    Me.txtStock.Text = ""
    Me.txtMensaje.Text = ""
    Me.txtMensaje.Enabled = False
    cmdBuscaTipo.Visible = True
End Sub

Private Sub txtMensaje_Change()
    If Me.optInsumo.value = True Then
            Call ValidaStr(txtMensaje)
    End If
    
End Sub

Private Sub txtStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    'If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Or KeyAscii = 46 Then KeyAscii = 0: Beep
    If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And (InStr(txtStock.Text, ".") Or optProducto.value) And (InStr(txtStock.Text, ".") Or optInsumo.value) Then KeyAscii = 0: Beep
End Sub
' Fin CGMiranda------------------------------------------------------
