VERSION 5.00
Begin VB.Form frmNuevoInvitado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Invitado"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frmNuevoInvitado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      Height          =   90
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   6675
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   4
      Left            =   4050
      Picture         =   "frmNuevoInvitado.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1500
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   3
      Left            =   5400
      Picture         =   "frmNuevoInvitado.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1500
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Apellidos"
      Height          =   555
      Index           =   2
      Left            =   5400
      TabIndex        =   3
      Top             =   660
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Nombres"
      Height          =   555
      Index           =   1
      Left            =   5400
      TabIndex        =   2
      Top             =   105
      Width           =   1275
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      MaxLength       =   149
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   5280
   End
   Begin VB.TextBox txtApellidos 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      MaxLength       =   149
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   630
      Width           =   5280
   End
End
Attribute VB_Name = "frmNuevoInvitado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public scodigoDelivery As String
Private Sub cmdOpcion_Click(Index As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If

   Select Case Index
         
          
          Case Is = 1 ' Razon social
               frmKeyBoard.txtResultado = txtNombre.Text
               frmKeyBoard.Show vbModal
               txtNombre.Text = IIf(wEnter, sDescrip, txtNombre.Text)
          
          Case Is = 2 ' Direccion
               frmKeyBoard.txtResultado = txtApellidos.Text
               frmKeyBoard.Show vbModal
               txtApellidos.Text = IIf(wEnter, sDescrip, txtApellidos.Text)
               
          Case Is = 3 ' Aceptar
               Dim nCorrela As String
               
               'Chequea Datos
               If txtNombre.Text = "" Then MsgBox "Ingrese el Nombre", vbExclamation, sMensaje: Exit Sub
               If txtApellidos = "" Then MsgBox "Ingrese el Apellido", vbExclamation, sMensaje: Exit Sub
            
               
               If frmBusquedaRapidaInvitado.wAdiciona Then
                  If Val(Calcular("select count(*) as Codigo from Tdeliveryinvitado where tNombre = '" & UCase(txtNombre.Text) & "' and tapellido='" & UCase(txtApellidos.Text) & "'", Cn)) > 0 Then
                     MsgBox "Invitado existente", vbCritical, sMensaje
                     Exit Sub
                  End If
               
                  'Obtiene el Correlativo
                  nCorrela = Calcular("select Max(tCodigoInvitado) as Codigo from Tdeliveryinvitado", Cn)
                
                  If IsNull(nCorrela) Or nCorrela = "" Then
                     sCodigo = "0000000001"
                  Else
                     sCodigo = Lib.Correlativo(nCorrela, 10)
                  End If
                  
                  Isql = "insert into tdeliveryInvitado( " & _
                         "tCodigoInvitado,tCodigoDelivery, tNombre, tApellido, tUsuario,  fRegistro) " & _
                         "values ('" & sCodigo & "', " & _
                                 " '" & scodigoDelivery & "', " & _
                                 " '" & UCase(txtNombre.Text) & "', " & _
                                 " '" & UCase(txtApellidos.Text) & "', " & _
                                 " '" & sUsuario & "',  getdate() )"
               Else
                  If Val(Calcular("select count(*) as Codigo from Tdeliveryinvitado where tNombre = '" & UCase(txtNombre.Text) & "' and tapellido='" & UCase(txtApellidos.Text) & "' and TCODIGOINVITADO<>'" & sCodigo & "' ", Cn)) > 0 Then
                     MsgBox "Invitado existente", vbCritical, sMensaje
                     Exit Sub
                  End If
               
                  Isql = "Update tdeliveryInvitado  SET " & _
                         "tNombre='" & UCase(txtNombre.Text) & "', " & _
                         "tApellido='" & UCase(txtApellidos.Text) & "', " & _
                         "fRegistro=getdate() " & _
                         "where tCodigoinvitado='" & sCodigo & "'"
               End If
               Cn.Execute Isql
               Unload Me
          
          Case Is = 4 ' Salir
               wEnter = False
               Unload Me
               
        
   End Select
End Sub

Private Sub Form_Load()
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   
 
   
   Centrar Me
   If frmBusquedaRapidaInvitado.wAdiciona Then
      Limpiar
    Else
      sCodigo = IIf(frmBusquedaRapidaInvitado.RsGrilla.EOF = True, "", frmBusquedaRapidaInvitado.RsGrilla!codigo)
      Mostrar
   End If
    
End Sub

Sub Mostrar()
    With frmBusquedaRapidaInvitado.RsGrilla
        txtNombre = IIf(IsNull(!tNombre), "", !tNombre)
        txtApellidos = IIf(IsNull(!tApellido), "", !tApellido)
    End With
End Sub

Sub Limpiar()
  Dim Control As Object
  For Each Control In Me.Controls
        If (TypeOf Control Is TextBox) Then
            Control.Text = ""
        End If
    Next Control
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmNuevoInvitado = Nothing
End Sub
