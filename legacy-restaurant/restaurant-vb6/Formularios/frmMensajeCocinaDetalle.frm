VERSION 5.00
Begin VB.Form frmMensajeCocinaDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2805
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9555
   Icon            =   "frmMensajeCocinaDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   9555
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
      TabIndex        =   15
      Top             =   0
      Width           =   7755
      Begin VB.CheckBox chkActivarAlerta 
         Alignment       =   1  'Right Justify
         Caption         =   "Activar Alerta"
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   1560
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
         Left            =   1185
         MaxLength       =   95
         TabIndex        =   1
         Text            =   " "
         Top             =   975
         Width           =   6420
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
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   390
         Width           =   1170
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   540
         TabIndex        =   2
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desc. Resumida :"
         Height          =   195
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje:"
         Height          =   195
         Left            =   420
         TabIndex        =   17
         Top             =   1020
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   480
         TabIndex        =   16
         Top             =   435
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9495
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2055
      Width           =   9555
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
         Left            =   7080
         Picture         =   "frmMensajeCocinaDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   8280
         Picture         =   "frmMensajeCocinaDetalle.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   4710
         Picture         =   "frmMensajeCocinaDetalle.frx":0636
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Left            =   5880
         Picture         =   "frmMensajeCocinaDetalle.frx":0B68
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4590
         TabIndex        =   14
         Top             =   60
         Width           =   4650
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmMensajeCocinaDetalle.frx":109A
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmMensajeCocinaDetalle.frx":15DC
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmMensajeCocinaDetalle.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4110
            Picture         =   "frmMensajeCocinaDetalle.frx":2060
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3630
            Picture         =   "frmMensajeCocinaDetalle.frx":25A2
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3120
            Picture         =   "frmMensajeCocinaDetalle.frx":2AE4
            Style           =   1  'Graphical
            TabIndex        =   10
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
            TabIndex        =   18
            Top             =   180
            Width           =   1665
         End
      End
   End
   Begin VB.Image Image 
      Height          =   2025
      Left            =   0
      Picture         =   "frmMensajeCocinaDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   30
      Width           =   1710
   End
End
Attribute VB_Name = "frmMensajeCocinaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CGMiranda-------------------------------------------------------------------
Option Explicit

Sub Asignar()
    With frmMensajeCocina.RsCabecera
        'Cuadro de Texto
        TxtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtMensaje = IIf(IsNull(!Mensaje), "", !Mensaje)
        
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
        chkActivarAlerta = IIf(!lActivarAlerta = True, 1, 0)
    End With
    
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmMensajeCocina.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmMensajeCocina.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmMensajeCocina.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmMensajeCocina.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmMensajeCocina.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmMensajeCocina.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmMensajeCocina.RsCabecera.AbsolutePosition & " de " & frmMensajeCocina.RsCabecera.RecordCount
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
               chkActivarAlerta.value = 0
               'Cambia el Nombre del Primer Text
               txtMensaje.SetFocus
                    
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               Dim nPos As Integer
             
               Set oComando = New clsComando
               'Chequea Datos
               If txtMensaje.Text = "" Then MsgBox "Ingrese el Mensaje ", vbExclamation, sMensaje: txtMensaje.SetFocus: Exit Sub
                                  
               If Sw Then
               
                If chkActivo.value = 1 Then
                    If Calcular("select count(Codigo) as codigo from TMENSAJECOCINA where lActivo=1", Cn) >= 30 Then
                        MsgBox "Solo puede tener Máximo 30 Mensajes a Cocina Activos", vbCritical, sMensaje
                        Exit Sub
                    End If
        
                End If
               
               
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(codigo) as Codigo from tmensajecocina where substring(codigo,1,2)= substring(ltrim(str(year(getdate()))),3,2) ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      nCorrela = Mid(Trim(str(Year(FechaServidor()))), 3, 2) & "000001"
                      TxtCodigo.Text = nCorrela
                  Else
                       nCorrela = Mid(Trim(str(Year(FechaServidor()))), 3, 2) & Lib.Correlativo(Mid(nCorrela, 3, 6), 6)
                       TxtCodigo.Text = nCorrela
                  End If
                  Sw = False
                txtMensaje.Text = UCase(txtMensaje.Text)
 
                'Llama al procedimiento para agregar un mensaje
                Isql = "usp_agregarmensaje"
                  If Not oComando.CreateCmdSp(Isql, Cn) Then
                        Set oComando = Nothing
                        Exit Sub
                    End If
                oComando.CreateParameter "@codigo", adVarChar, adParamInput, 8, TxtCodigo.Text
                oComando.CreateParameter "@usuario", adVarChar, adParamInput, 15, sUsuario
                oComando.CreateParameter "@mensaje", adVarChar, adParamInput, 100, txtMensaje.Text
                oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 3, sCaja
                oComando.CreateParameter "@lactivo", adBoolean, adParamInput, 1, chkActivo.value
                oComando.CreateParameter "@lactivaralerta", adBoolean, adParamInput, 1, chkActivarAlerta.value
                  If Not oComando.GetParamOK Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  If Not oComando.ExecSP Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  frmMensajeCocina.RsCabecera.Filter = ""
                  frmMensajeCocina.RsCabecera.Sort = "Codigo ASC"
                  frmMensajeCocina.RsCabecera.Requery
                  frmMensajeCocina.RsCabecera.MoveLast
                  ActivarBotones (True)
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  cmdTexto.Caption = "Registro " & IIf(frmMensajeCocina.RsCabecera.RecordCount = 0, 0, frmMensajeCocina.RsCabecera.AbsolutePosition) & " de " & frmMensajeCocina.RsCabecera.RecordCount

               Else
                    txtMensaje.Text = UCase(txtMensaje.Text)
                    
                    If chkActivo.value = 1 Then
                    If Calcular("select count(Codigo) as codigo from TMENSAJECOCINA where lActivo=1 and codigo <>'" & TxtCodigo.Text & "'", Cn) >= 30 Then
                        MsgBox "Solo puede tener Máximo 30 Mensajes a Cocina Activos", vbCritical, sMensaje
                        Exit Sub
                    End If
        
                    End If
                    
                    Isql = "usp_modificarmensaje"
                    If Not oComando.CreateCmdSp(Isql, Cn) Then
                        Set oComando = Nothing
                        Exit Sub
                    End If
                    oComando.CreateParameter "@usuario", adVarChar, adParamInput, 15, sUsuario
                    oComando.CreateParameter "@codigo", adVarChar, adParamInput, 8, TxtCodigo.Text
                    oComando.CreateParameter "@mensaje", adVarChar, adParamInput, 100, txtMensaje.Text
                    oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 3, sCaja
                    oComando.CreateParameter "@lactivo", adBoolean, adParamInput, 1, chkActivo.value
                    oComando.CreateParameter "@lactivaralerta", adBoolean, adParamInput, 1, chkActivarAlerta.value
                    If Not oComando.GetParamOK Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  If Not oComando.ExecSP Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                   nPos = frmMensajeCocina.RsCabecera.AbsolutePosition
                   frmMensajeCocina.RsCabecera.Filter = ""
                   frmMensajeCocina.RsCabecera.Requery
                   frmMensajeCocina.RsCabecera.AbsolutePosition = nPos
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          Case Is = 2 '
                    Set oComando = New clsComando
                    
                If frmMensajeCocina.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
                              'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Mensaje?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
                    Isql = "USP_ELIMINARRMENSAJES"
                    If Not oComando.CreateCmdSp(Isql, Cn) Then
                        Set oComando = Nothing
                        Exit Sub
                    End If
                    oComando.CreateParameter "@codigo", adVarChar, adParamInput, 8, TxtCodigo.Text
                   
                    If Not oComando.GetParamOK Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                  If Not oComando.ExecSP Then
                     Set oComando = Nothing
                     Exit Sub
                  End If
                    frmMensajeCocina.RsCabecera.Requery
                    If frmMensajeCocina.RsCabecera.RecordCount <> 0 Then
                            frmMensajeCocina.RsCabecera.MoveLast
                            Asignar
                            cmdTexto.Caption = "Registro " & IIf(frmMensajeCocina.RsCabecera.RecordCount = 0, 0, frmMensajeCocina.RsCabecera.AbsolutePosition) & " de " & frmMensajeCocina.RsCabecera.RecordCount
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
    Me.Caption = " Mantenimiento de Mensajes"
    fraDetalle.Caption = Me.Caption
    
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.value = 1
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & frmMensajeCocina.RsCabecera.AbsolutePosition & " de " & frmMensajeCocina.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmMensajeCocinaDetalle = Nothing
End Sub

Sub ActivarBotones(ByVal Activa As Boolean)
    cmdNavegar(0).Enabled = Activa
    cmdNavegar(1).Enabled = Activa
    cmdNavegar(2).Enabled = Activa
    cmdNavegar(3).Enabled = Activa
    cmdNavegar(4).Enabled = Activa
    cmdNavegar(5).Enabled = Activa
    cmdOpcion(0).Enabled = Activa
End Sub

Private Sub txtMensaje_Change()
    Call ValidaStr(txtMensaje)
End Sub
' Fin CGMiranda------------------------------------------------------
