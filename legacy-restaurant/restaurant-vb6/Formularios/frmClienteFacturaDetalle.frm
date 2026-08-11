VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmClienteFacturaDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6420
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9870
   Icon            =   "frmClienteFacturaDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   9870
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
      Height          =   5655
      Left            =   1680
      TabIndex        =   19
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton Command2 
         Caption         =   "Agente Retencion"
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
         Left            =   5200
         TabIndex        =   48
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkAgenteRetencion 
         Caption         =   "Es Agente de retención?"
         Height          =   255
         Left            =   1530
         TabIndex        =   47
         Top             =   1840
         Width           =   2175
      End
      Begin VB.Frame FrmDatosSeparados 
         Caption         =   "Datos Separados"
         Enabled         =   0   'False
         Height          =   1815
         Left            =   120
         TabIndex        =   38
         Top             =   3720
         Width           =   7815
         Begin VB.TextBox txtSegundoApellido 
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
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   44
            Text            =   " "
            Top             =   1320
            Width           =   6090
         End
         Begin VB.TextBox txtPrimerApellido 
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
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   43
            Text            =   " "
            Top             =   960
            Width           =   6090
         End
         Begin VB.TextBox txtSegundoNombre 
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
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   42
            Text            =   " "
            Top             =   600
            Width           =   6090
         End
         Begin VB.TextBox txtPrimerNombre 
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
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   40
            Text            =   " "
            Top             =   240
            Width           =   6090
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Segundo Apellido:"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   1320
            Width           =   1290
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Primer Apellido:"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   960
            Width           =   1080
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Segundo Nombre:"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   600
            Width           =   1290
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Primer Nombre:"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   1080
         End
      End
      Begin VB.CommandButton cmdVerificaDNI 
         Caption         =   "Verificar DNI"
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
         Left            =   3840
         TabIndex        =   37
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtUrbanizacion 
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
         Left            =   1515
         MaxLength       =   200
         TabIndex        =   34
         Text            =   " "
         Top             =   2160
         Width           =   6210
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ubigeo"
         Height          =   300
         Left            =   5640
         TabIndex        =   33
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtCodigoUbigeo 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6495
         MaxLength       =   50
         TabIndex        =   32
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdVerifica 
         Caption         =   "Verificar Ruc"
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
         Left            =   6480
         TabIndex        =   31
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtEnlace 
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
         Left            =   4890
         MaxLength       =   50
         TabIndex        =   28
         Text            =   " "
         Top             =   3285
         Width           =   2835
      End
      Begin VB.TextBox txtCorreo 
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
         Left            =   1515
         MaxLength       =   200
         TabIndex        =   5
         Text            =   " "
         Top             =   2925
         Width           =   6210
      End
      Begin VB.TextBox txtDireccion 
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
         Left            =   1515
         MaxLength       =   200
         TabIndex        =   4
         Text            =   " "
         Top             =   2535
         Width           =   6210
      End
      Begin VB.TextBox txtDetallado 
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
         Left            =   1515
         MaxLength       =   200
         TabIndex        =   1
         Text            =   " "
         Top             =   753
         Width           =   6210
      End
      Begin VB.TextBox txtResumido 
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
         Left            =   4890
         MaxLength       =   15
         TabIndex        =   3
         Text            =   " "
         Top             =   1116
         Width           =   2835
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
         Left            =   1515
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
         Left            =   840
         TabIndex        =   6
         Top             =   3330
         Width           =   840
      End
      Begin MSDataListLib.DataCombo cboTipoIdentidad 
         Height          =   315
         Left            =   1515
         TabIndex        =   2
         Top             =   1080
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboTipoCliente 
         Height          =   315
         Left            =   1515
         TabIndex        =   30
         Top             =   1485
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cboTipoResponsable 
         Height          =   315
         Left            =   2880
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Urbanizacion :"
         Height          =   195
         Left            =   420
         TabIndex        =   35
         Top             =   2175
         Width           =   1020
      End
      Begin VB.Label lblTipoCliente 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cliente :"
         Height          =   195
         Left            =   510
         TabIndex        =   29
         Top             =   1530
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Enlace"
         Height          =   195
         Left            =   4200
         TabIndex        =   27
         Top             =   3330
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Identificación :"
         Height          =   195
         Left            =   45
         TabIndex        =   26
         Top             =   1125
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Correo :"
         Height          =   195
         Left            =   885
         TabIndex        =   25
         Top             =   2925
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dirección :"
         Height          =   195
         Left            =   675
         TabIndex        =   24
         Top             =   2535
         Width           =   765
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Razón Social :"
         Height          =   195
         Left            =   405
         TabIndex        =   22
         Top             =   780
         Width           =   1035
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Identificador :"
         Height          =   195
         Left            =   3960
         TabIndex        =   21
         Top             =   1125
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   855
         TabIndex        =   20
         Top             =   435
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9810
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5670
      Width           =   9870
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
         Left            =   8220
         Picture         =   "frmClienteFacturaDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   60
         Width           =   1170
      End
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
         Left            =   7050
         Picture         =   "frmClienteFacturaDetalle.frx":0534
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "frmClienteFacturaDetalle.frx":0636
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmClienteFacturaDetalle.frx":0B68
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4590
         TabIndex        =   12
         Top             =   60
         Width           =   4650
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmClienteFacturaDetalle.frx":109A
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmClienteFacturaDetalle.frx":15DC
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmClienteFacturaDetalle.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4110
            Picture         =   "frmClienteFacturaDetalle.frx":2060
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3630
            Picture         =   "frmClienteFacturaDetalle.frx":25A2
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3150
            Picture         =   "frmClienteFacturaDetalle.frx":2AE4
            Style           =   1  'Graphical
            TabIndex        =   13
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
            TabIndex        =   23
            Top             =   180
            Width           =   1665
         End
      End
   End
   Begin VB.Image Image 
      Height          =   5745
      Left            =   0
      Picture         =   "frmClienteFacturaDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   30
      Width           =   1710
   End
End
Attribute VB_Name = "frmClienteFacturaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTipoIdentidad As Recordset
Dim RsTipoResponsable As Recordset
Dim RsTipoCliente As Recordset
Dim RsParametro As Recordset
'Dim Isql As String
Public lagregar As Boolean
Dim lDatoSeparados As Boolean

Sub Asignar()
    With frmClienteFactura.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tIdentidad), "", !tIdentidad)
        txtDireccion = IIf(IsNull(!tDireccion), "", !tDireccion)
        
        txtCorreo = IIf(IsNull(!tcorreo), "", !tcorreo)

        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
        txtEnlace = IIf(IsNull(!tEnlace), "", !tEnlace)
        chkAgenteRetencion = IIf(!AgenteRetencion = True, 1, 0)
        Me.txtPrimerNombre = IIf(IsNull(!tNombre), "", !tNombre)
        Me.txtSegundoNombre = IIf(IsNull(!tNombre2), "", !tNombre2)
        Me.txtPrimerApellido = IIf(IsNull(!tApellido), "", !tApellido)
        Me.txtSegundoApellido = IIf(IsNull(!tApellido2), "", !tApellido2)
        
        cboTipoIdentidad.BoundText = IIf(IsNull(!tTipoIdentidad), "", !tTipoIdentidad)
        'Tipo cliente
        cboTipoCliente.BoundText = IIf(IsNull(!tTipoCliente), "", !tTipoCliente)
        
        Me.txtCodigoUbigeo = IIf(IsNull(!CodigoUbigeo), "", !CodigoUbigeo)
        Me.txtUrbanizacion = IIf(IsNull(!urbanizacion), "", !urbanizacion)
        Me.cboTipoResponsable.BoundText = IIf(IsNull(!ttiporesponsable), "", !ttiporesponsable)

        
        If (Calcular("Select isnull(lActivaDatosSeparados,0) as codigo from vtipoidentidad where Codigo='" & Me.cboTipoIdentidad.BoundText & "'", Cn)) Then
            lDatoSeparados = True
        Else
            lDatoSeparados = False
        End If
        
    End With
    
End Sub



Private Sub cboTipoIdentidad_Change()
On Error GoTo fin
    If (Calcular("Select isnull(lActivaDatosSeparados,0) as codigo from vtipoidentidad where Codigo='" & Me.cboTipoIdentidad.BoundText & "'", Cn)) Then
        FrmDatosSeparados.Enabled = True
        FrmDatosSeparados.backColor = &HC0C0FF
        Me.txtDetallado.Enabled = False
        lDatoSeparados = True
    Else
        FrmDatosSeparados.Enabled = False
        FrmDatosSeparados.backColor = &H8000000F
        Me.txtDetallado.Enabled = True
        lDatoSeparados = False
    End If
    Exit Sub
fin:
    MsgBox "Error: " & Error
End Sub


Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmClienteFactura.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmClienteFactura.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmClienteFactura.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmClienteFactura.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmClienteFactura.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmClienteFactura.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmClienteFactura.RsCabecera.AbsolutePosition & " de " & frmClienteFactura.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)

   Dim xtTipoIdentidad As String
   Dim Numero As String
   Dim correo As String

   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               lagregar = True
               LlenaCombos
               cboTipoIdentidad.BoundText = ""
               chkActivo.Value = 1
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
                    
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               Dim nPos As Variant
               Me.txtResumido.Text = Trim(Me.txtResumido.Text)
               'Chequea Datos
                If (Calcular("Select isnull(lActivaDatosSeparados,0) as codigo from vtipoidentidad where Codigo='" & Me.cboTipoIdentidad.BoundText & "'", Cn)) Then
                    If Trim(Me.txtPrimerNombre.Text) = "" Then MsgBox "Ingrese Primer Nombre", vbExclamation, sMensaje: Me.txtPrimerNombre.SetFocus: Exit Sub
                    If Trim(Me.txtSegundoNombre.Text) = "" Then MsgBox "Ingrese Segundo Nombre", vbExclamation, sMensaje: Me.txtSegundoNombre.SetFocus: Exit Sub
                    If Trim(Me.txtPrimerApellido.Text) = "" Then MsgBox "Ingrese Primer Apellido", vbExclamation, sMensaje: Me.txtPrimerApellido.SetFocus: Exit Sub
                    If Trim(Me.txtSegundoApellido.Text) = "" Then MsgBox "Ingrese Segundo Apellido", vbExclamation, sMensaje: Me.txtSegundoApellido.SetFocus: Exit Sub
                Else
                    If Trim(txtDetallado.Text) = "" Then MsgBox "Ingrese la Razón Social", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
                End If
    
               'If Trim(txtDetallado.Text) = "" Then MsgBox "Ingrese la Razón Social", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese el Id. Tributario", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
               If cboTipoIdentidad.Text = "" Then MsgBox "Seleccione el Tipo de Identidad", vbExclamation, sMensaje: cboTipoIdentidad.SetFocus: Exit Sub
               
               If lSAP Then
                    If cboTipoCliente.Text = "" Then MsgBox "Seleccione el Tipo de Cliente.", vbExclamation, sMensaje: cboTipoCliente.SetFocus: Exit Sub
               End If
               
               If pais = "003" Then
                    If cboTipoResponsable.Text = "" Then MsgBox "Seleccione el Tipo Responsable.", vbExclamation, sMensaje: cboTipoResponsable.SetFocus: Exit Sub
                    If Trim(txtDireccion.Text) = "" Then MsgBox "Ingrese la direccion del Cliente", vbExclamation, sMensaje: Me.txtDireccion.SetFocus: Exit Sub
               End If
               
               ' cambios para validar DNI y numeros
               If RsParametro!lValidaDNI = True And pais = "002" Then
                  If cboTipoIdentidad.SelectedItem = 2 Then
                    Numero = modProcedimiento.ValidarDNI(LTrim(Me.txtResumido))
                    If Numero = False Then
                    MsgBox "El DNI ingresado no es valido", vbCritical, sMensaje
                    Exit Sub
                    End If
                  End If
                  ' validar correo
                 correo = modProcedimiento.Validar_Email(Me.txtCorreo.Text)
                 If correo = False Then
                    MsgBox "El correo electronico es invalido", vbCritical, sMensaje
                 Exit Sub
                 End If
                  
               End If
               '---------------------------------
                    
               If pais = "002" Then ' ECUADOR
'                        If Calcular("Select isnull(nValor,0) As Codigo from vtipoidentidad where Codigo= '" & Me.cboTipoIdentidad.BoundText & "'", Cn) Then
'                            If Len(Trim(txtResumido.Text)) = 13 Or Len(Trim(txtResumido.Text)) = 10 Then
'
'                            Else
'                               MsgBox "La longitud del Identificador debe ser 10(Cédula) ó 13(RUC)", vbCritical, sMensaje
'                               Exit Sub
'                            End If
'                        End If
'
'                        If Len(Trim(txtResumido.Text)) = 10 Then
'                            xtTipoIdentidad = "01"
'                        ElseIf Len(Trim(txtResumido.Text)) = 13 Then
'                            xtTipoIdentidad = "02"
'                        End If
                ElseIf pais = "003" Then ' ARGENTINA
                        If Calcular("Select isnull(nValor,0) As Codigo from vtipoidentidad where Codigo= '" & Me.cboTipoIdentidad.BoundText & "'", Cn) Then
                                If lLongitud And Len(Trim(txtResumido.Text)) <> nLongitud Then
                                   MsgBox "La longitud del Identificador debe ser " & nLongitud, vbCritical, sMensaje
                                   Exit Sub
                                ElseIf Not lLongitud And Len(Trim(txtResumido.Text)) < nLongitud Then
                                   MsgBox "La longitud del Identificador debe ser mayor igual a " & nLongitud, vbCritical, sMensaje
                                   Exit Sub
                                End If
                                Dim ValTemp As String
                                ValTemp = ValidaCUIT(txtResumido.Text)
                                If ValTemp = "" Then
                                   'MsgBox "El número ingresado no es válido", vbCritical, sMensaje
                                   Exit Sub
                                Else
                                    Me.txtResumido.Text = ValTemp
                                End If
                        End If
                
                Else
'                        'PERU - BOLIVIA
'                        If Calcular("Select isnull(nValor,0) As Codigo from vtipoidentidad where Codigo= '" & Me.cboTipoIdentidad.BoundText & "'", Cn) Then
'
'                                If lLongitud And Len(Trim(txtResumido.Text)) <> nLongitud Then
'                                   MsgBox "La longitud del Identificador debe ser " & nLongitud, vbCritical, sMensaje
'                                   Exit Sub
'                                ElseIf Not lLongitud And Len(Trim(txtResumido.Text)) < nLongitud Then
'                                   MsgBox "La longitud del Identificador debe ser mayor igual a " & nLongitud, vbCritical, sMensaje
'                                   Exit Sub
'                                End If
'
'                                If Not ValidaRuc(txtResumido.Text) Then
'                                   MsgBox "El número ingresado no es válido", vbCritical, sMensaje
'                                   Exit Sub
'                                End If
'                        End If
'                        xtTipoIdentidad = ""


                    xtTipoIdentidad = ""
                    
               End If
               
                Dim RsValidacionDatos As Recordset
                Set RsValidacionDatos = Lib.OpenRecordset("exec  usp_ValidaTipoIdentidad '" & cboTipoIdentidad.BoundText & "','" & txtResumido.Text & "'", Cn)
                If RsValidacionDatos.RecordCount > 0 Then
                    RsValidacionDatos.MoveFirst
                    If Left(RsValidacionDatos!respuesta, 1) = "1" Then
                        MsgBox "Validacion: " & RsValidacionDatos!respuesta, vbInformation
                        Exit Sub
                    End If
                End If
                
                If Trim(cboTipoIdentidad.Text) = "RUC" And pais = "000" Then
                    If Not ValidaRuc(Trim(txtResumido.Text)) Then
                       MsgBox "El número RUC ingresado no es válido", vbCritical, sMensaje
                       txtResumido.Text = ""
                       wEnter = False
                       Exit Sub
                    End If
                End If
                    
               If lFEpape And (cboTipoIdentidad.BoundText = "01" Or cboTipoIdentidad.BoundText = "02") Then
                    If txtCorreo.Text = "" Then MsgBox "Ingrese el Correo Electrónico", vbExclamation, sMensaje: txtCorreo.SetFocus: Exit Sub
                    'VALIDA MAIL
                    If Not Validar_Email(txtCorreo.Text) Then
                       MsgBox "El Correo ingresado no es válido", vbCritical, sMensaje
                       Exit Sub
                    End If
               End If

               If val(Calcular("select tIdentidad as Codigo from TCLIENTE where tIdentidad = '" & Trim(txtResumido.Text) & "'", Cn)) > 0 And Sw Then
                  MsgBox "Identificador Repetido", vbCritical, sMensaje
                  Exit Sub
               End If
                    
               If Sw Then
                  If Calcular("select count(tIdentidad) as Codigo from TCLIENTE where tIdentidad='" & Trim(txtResumido.Text) & "'", Cn) > 0 Then
                     MsgBox "Error: Identificador Existente", vbCritical, sMensaje
                     Exit Sub
                  End If
                              
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigoCliente) as Codigo from TCLIENTE", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "00001"
                  Else
                    'AJMM 27-09-2025
                    Dim letras As String
                    Dim numeroCor As Long
                    Dim i As Integer
                    
                    letras = ""
                    numeroCor = 0
                    
                    ' Extraemos las letras del inicio
                    i = 1
                    Do While i <= Len(nCorrela) And IsLetter(Mid$(nCorrela, i, 1))
                        letras = letras & Mid$(nCorrela, i, 1)
                        i = i + 1
                    Loop
                    
                    ' Extraemos la parte numérica
                    If i <= Len(nCorrela) Then
                        numeroCor = CLng(Mid$(nCorrela, i))
                    End If
                    
                    ' Incrementamos número
                    numeroCor = numeroCor + 1
                    
                    ' Cantidad de dígitos que deben quedar para número
                    Dim numDig As Integer
                    numDig = 5 - Len(letras)
                    
                    ' Si el número se desborda (ej: 99999 ? 100000), incrementamos letras
                    If Len(CStr(numeroCor)) > numDig Then
                        numeroCor = 1 ' reiniciamos numeración
                        letras = IncrementarLetras(letras)
                        numDig = 5 - Len(letras)
                    End If
                    
                    ' Formamos el código final (siempre 5 caracteres)
                    txtCodigo.Text = letras & Right$(String(numDig, "0") & CStr(numeroCor), numDig)
'                      txtCodigo.Text = Lib.Correlativo(nCorrela, 5)
'AJMM 20-02-2025
'                    Dim letras As String
'                    Dim numeroCor As Long
'                    Dim numeroLetra As Long
'                    Dim letra1 As String
'                    Dim letra2 As String
'                    Dim letra3 As String
'                    Dim letra4 As String
'                    Dim letra5 As String
'
'                    ' Inicializamos las variables
'                        letras = ""
'                        numeroCor = 0
'                        letra1 = ""
'                        letra2 = ""
'                        letra3 = ""
'                        letra4 = ""
'                        letra5 = ""
'
'                        ' Separar las letras y los números
'                        Dim i As Integer
'                        i = 1
'                        ' Extraemos las letras al principio
'                        Do While i <= Len(nCorrela) And IsLetter(Mid(nCorrela, i, 1))
'                            letras = letras & Mid(nCorrela, i, 1)
'                            i = i + 1
'                        Loop
'
'                        ' Extraemos la parte numérica
'                        If i <= Len(nCorrela) Then
'                            numeroCor = CLng(Mid(nCorrela, i))
'                        End If
                    
'                        ' Si el número es menor a 99999, incrementamos el número
'                        If numeroCor < 99999 And numeroCor < 9999 Then
'                            numeroCor = numeroCor + 1
'                            If numeroCor < 10 Then
'                                txtCodigo.Text = letras & Right("000" & CStr(numeroCor), 5)
'                            End If
'
'                            If numeroCor >= 10 And numeroCor < 100 Then
'                                txtCodigo.Text = letras & Right("00" & CStr(numeroCor), 5)
'                            End If
'
'                            If numeroCor >= 100 And numeroCor < 1000 Then
'                                txtCodigo.Text = letras & Right("0" & CStr(numeroCor), 5)
'                            End If
'
'                            If numeroCor >= 1000 And numeroCor < 10000 Then
'                                txtCodigo.Text = letras & Right("" & CStr(numeroCor), 5)
'                            End If
'
'                        Else
'                            ' Si llegamos a 99999, incrementamos las letras
'                            If Len(letras) = 1 Then
'                                ' Solo una letra, incrementamos la siguiente letra
'                                letra1 = Chr(Asc(letras) + 1)
'                                txtCodigo.Text = letra1 & "0001"
'                            ElseIf Len(letras) = 2 Then
'                                ' Dos letras, incrementamos la segunda letra
'                                letra1 = Mid(letras, 1, 1)
'                                letra2 = Mid(letras, 2, 1)
'
'                                If letra2 = "Z" Then
'                                    letra2 = "A"
'                                    letra1 = Chr(Asc(letra1) + 1)
'                                Else
'                                    letra2 = Chr(Asc(letra2) + 1)
'                                End If
'
'                                txtCodigo.Text = letra1 & letra2 & "001"
'                            Else
'                                ' Aquí se puede agregar más lógica si llegamos a más de 2 letras
'                                ' (Por ejemplo, "AAA00001")
'                                txtCodigo.Text = "A0001" ' Ejemplo de próxima extensión
'                            End If
'                        End If
                         ' Si el número es menor a 9999, incrementamos el número

'                    If numeroCor < 9999 Then
'                        numeroCor = numeroCor + 1
'                        txtCodigo.Text = letras & Right$("0000" & CStr(numeroCor), 4) ' Aseguramos que el número tenga 4 dígitos
'                    Else
'                        ' Si llegamos a 9999, incrementamos las letras
'                        If Len(letras) = 1 Then
'                            ' Solo una letra, incrementamos la siguiente letra
'                            letra1 = Chr(Asc(letras) + 1)
'                            If Asc(letra1) > 90 Then ' Si la letra es más allá de "Z"
'                                letra1 = "A" ' Reiniciamos a "A"
'                            End If
'                            txtCodigo.Text = letra1 & "A001"
'                        ElseIf Len(letras) = 2 Then
'                            ' Dos letras, incrementamos la segunda letra
'                            letra1 = Mid$(letras, 1, 1)
'                            letra2 = Mid$(letras, 2, 1)
'
'                            If letra2 = "Z" Then
'                                letra2 = "A"
'                                letra1 = Chr(Asc(letra1) + 1)
'                                If Asc(letra1) > 90 Then
'                                    letra1 = "A" ' Reiniciamos la primera letra si llega a "Z"
'                                End If
'                            Else
'                                letra2 = Chr(Asc(letra2) + 1)
'                            End If
'
'                            txtCodigo.Text = letra1 & letra2 & "0001"
'                        ElseIf Len(letras) = 3 Then
'                            ' Tres letras, incrementamos la tercera letra
'                            letra1 = Mid$(letras, 1, 1)
'                            letra2 = Mid$(letras, 2, 1)
'                            letra3 = Mid$(letras, 3, 1)
'
'                            If letra3 = "Z" Then
'                                letra3 = "A"
'                                If letra2 = "Z" Then
'                                    letra2 = "A"
'                                    If letra1 = "Z" Then
'                                        letra1 = "A"
'                                    Else
'                                        letra1 = Chr(Asc(letra1) + 1)
'                                    End If
'                                Else
'                                    letra2 = Chr(Asc(letra2) + 1)
'                                End If
'                            Else
'                                letra3 = Chr(Asc(letra3) + 1)
'                            End If
'
'                            txtCodigo.Text = letra1 & letra2 & letra3 & "0001"
'                        ElseIf Len(letras) = 4 Then
'                            ' Cuatro letras, incrementamos la cuarta letra
'                            letra1 = Mid$(letras, 1, 1)
'                            letra2 = Mid$(letras, 2, 1)
'                            letra3 = Mid$(letras, 3, 1)
'                            letra4 = Mid$(letras, 4, 1)
'
'                            If letra4 = "Z" Then
'                                letra4 = "A"
'                                If letra3 = "Z" Then
'                                    letra3 = "A"
'                                    If letra2 = "Z" Then
'                                        letra2 = "A"
'                                        If letra1 = "Z" Then
'                                            letra1 = "A"
'                                        Else
'                                            letra1 = Chr(Asc(letra1) + 1)
'                                        End If
'                                    Else
'                                        letra2 = Chr(Asc(letra2) + 1)
'                                    End If
'                                Else
'                                    letra3 = Chr(Asc(letra3) + 1)
'                                End If
'                            Else
'                                letra4 = Chr(Asc(letra4) + 1)
'                            End If
'
'                            txtCodigo.Text = letra1 & letra2 & letra3 & letra4 & "0001"
'                        ElseIf Len(letras) = 5 Then
'                            ' Cinco letras, ya no se deben incrementar más letras
'                            ' Si llegamos a 5 letras, aumentamos el número y volvemos a "A" si es necesario
'                            txtCodigo.Text = "AAAA1"
'                        ElseIf numeroCor < 99999 Then
'                            txtCodigo.Text = Lib.Correlativo(nCorrela, 5)
'                        Else
'                            txtCodigo.Text = "A0001"
'                        End If
'                    End If
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
                  Isql = "insert into TCLIENTE( " & _
                         "tCodigoCliente, tEmpresa, tIdentidad, tDireccion, tUsuario, fRegistro, tCorreo, tEnlace, tTipoIdentidad,tTipoCliente, AgenteRetencion, lActivo, tUbigeo, tUrbanizacion, tTipoResponsable,tNombre, tNombre2, tApellido, tApellido2) " & _
                         "values ( '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                " '" & txtDireccion.Text & "', " & _
                                " '" & sUsuario & "', getdate(), " & _
                                " '" & txtCorreo.Text & "', " & _
                                " '" & txtEnlace.Text & "', " & _
                                " '" & cboTipoIdentidad.BoundText & "', " & _
                                " '" & cboTipoCliente.BoundText & "', " & _
                                " '" & chkAgenteRetencion.Value & "'," & _
                                       chkActivo.Value & ",'" & Me.txtCodigoUbigeo.Text & "', '" & Me.txtUrbanizacion.Text & "','" & Me.cboTipoResponsable.BoundText & "', '" & Trim(Me.txtPrimerNombre.Text) & "', '" & Trim(Me.txtSegundoNombre.Text) & "', '" & Trim(Me.txtPrimerApellido.Text) & "', '" & Trim(Me.txtSegundoApellido.Text) & "')"
            
                  Cn.Execute Isql
                  frmClienteFactura.RsCabecera.Sort = "Codigo ASC"
                  frmClienteFactura.RsCabecera.Requery
                  frmClienteFactura.RsCabecera.MoveLast
                  ActivarBotones (True)
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  cmdTexto.Caption = "Registro " & IIf(frmClienteFactura.RsCabecera.RecordCount = 0, 0, frmClienteFactura.RsCabecera.AbsolutePosition) & " de " & frmClienteFactura.RsCabecera.RecordCount
               
               Else
               
                  If lFEpape And (cboTipoIdentidad.BoundText = "01" Or cboTipoIdentidad.BoundText = "02") Then
                         If txtCorreo.Text = "" Then MsgBox "Ingrese el Correo Electrónico", vbExclamation, sMensaje: txtCorreo.SetFocus: Exit Sub
                          'VALIDA MAIL
                         If Not Validar_Email(txtCorreo.Text) Then
                            MsgBox "El Correo ingresado no es válido", vbCritical, sMensaje
                            Exit Sub
                         End If
                  End If
                    
                  'Cambiar el SQL
                  If Calcular("select count(tIdentidad) as Codigo from TCLIENTE where tCodigoCliente <>'" & txtCodigo.Text & "' and tIdentidad='" & txtResumido.Text & "'", Cn) > 0 Then
                     MsgBox "Error: Identificador Existente", vbCritical, sMensaje
                     Exit Sub
                  End If
                  
                  Isql = "update TCLIENTE set " & _
                         "tEmpresa ='" & txtDetallado.Text & "', " & _
                         "tIdentidad ='" & txtResumido.Text & "', " & _
                         "tDireccion ='" & txtDireccion.Text & "', " & _
                         "tCorreo ='" & txtCorreo.Text & "', " & _
                         "tEnlace ='" & txtEnlace.Text & "', " & _
                         "tTipoIdentidad = '" & cboTipoIdentidad.BoundText & "', " & _
                         "tTipoCliente = '" & cboTipoCliente.BoundText & "', " & _
                         "lActivo =" & chkActivo.Value & "," & _
                         "AgenteRetencion =" & chkAgenteRetencion.Value & "," & _
                         "tUbigeo ='" & Me.txtCodigoUbigeo.Text & "'," & _
                         "tUrbanizacion ='" & Me.txtUrbanizacion.Text & "'," & _
                         "tTipoResponsable ='" & Me.cboTipoResponsable.BoundText & "'," & _
                         "tNombre ='" & Trim(Me.txtPrimerNombre.Text) & "'," & _
                         "tNombre2 ='" & Trim(Me.txtSegundoNombre.Text) & "'," & _
                         "tApellido ='" & Trim(Me.txtPrimerApellido.Text) & "'," & _
                         "tApellido2 ='" & Trim(Me.txtSegundoApellido.Text) & "'," & _
                         "lreplica=1 where tCodigoCliente = '" & txtCodigo & "'"

                   Cn.Execute Isql
                   nPos = frmClienteFactura.RsCabecera.Bookmark
                   frmClienteFactura.RsCabecera.Requery
                   If frmClienteFactura.RsCabecera.RecordCount = 0 Then
                      frmClienteFactura.RsCabecera.Filter = adFilterNone
                   End If
                   frmClienteFactura.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               
            Cn.Execute ("update tcliente set lnuevomodificado=1 where tCodigoCliente='" & txtCodigo.Text & "'")
          
          Case Is = 2 ' Eliminar
               If frmClienteFactura.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Cliente" & txtDetallado.Text & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               If Calcular("select count(tDocumento) as Codigo from MDOCUMENTO where tCodigoCliente ='" & txtCodigo.Text & "'", Cn) > 0 Then
                  MsgBox "Error: Se han generado documentos con este Cliente" & Chr(13) & "No se puede Eliminar", vbCritical, sMensaje
                  Exit Sub
               End If
                               
               'Cambia el Delete
               Cn.Execute "delete from TCLIENTE where tCodigoCliente = '" & txtCodigo & "'"
               frmClienteFactura.RsCabecera.Requery
               If frmClienteFactura.RsCabecera.RecordCount <> 0 Then
                  frmClienteFactura.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmClienteFactura.RsCabecera.RecordCount = 0, 0, frmClienteFactura.RsCabecera.AbsolutePosition) & " de " & frmClienteFactura.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If
          
          Case Is = 3 ' Salir
               Unload Me
   End Select

End Sub

Private Sub cmdVerifica_Click()
On Error GoTo fin
    Screen.MousePointer = vbHourglass
    Dim xDat As String
    Dim xRazSoc As String, xEst As String, xCon As String, xDir As String
    Dim xRazSocX As Long, xEstX As Long, xConX As Long, xDirX As Long
    Dim xRazSocY As Long, xEstY As Long, xConY As Long, xDirY As Long
    Dim xWml As New XMLHTTP
    
     Me.txtPrimerNombre.Text = ""
    
    If Not IsNumeric(Trim(txtResumido.Text)) Then
        MsgBox "formato incorrectO, Verificar!!!", vbInformation, sMensaje
        Exit Sub
    End If
    If Len(Trim(txtResumido.Text)) <> 11 Then
        MsgBox "longitud incorrecto, Verificar!!!", vbInformation, sMensaje
        Exit Sub
    End If

    xWml.Open "POST", RutaConsultaRuc & "/" & Trim(txtResumido.Text), False
    xWml.send
    If xWml.Status = 200 Then
        xDat = Trim(xWml.responseText)
        xDat = Replace(xDat, vbNewLine, " ")
        xDat = Replace(Replace(Replace(xDat, "," + Chr(34), vbNewLine), "{", ""), "}", "")
        xDat = Replace(xDat, Chr(34), "")
        
        Dim MATRIZ() As String
        MATRIZ = Split(xDat, vbNewLine)
        
        MATRIZ(0) = Replace(MATRIZ(0), "status:", "")
        If (MATRIZ(0) = 1) Then
            MATRIZ(2) = Replace(MATRIZ(2), "ruc:", "")
            MATRIZ(3) = Replace(MATRIZ(3), "razonSocial:", "")
            MATRIZ(4) = Replace(MATRIZ(4), "direccion:", "")
            MATRIZ(5) = Replace(MATRIZ(5), "ubigeo:", "")
            
            Me.txtUrbanizacion.Text = Calcular("select tdistrito as codigo from tubigeo where tcodigo='" & MATRIZ(5) & "'", Cn)
            Me.txtDetallado.Text = MATRIZ(3)

            
            If lDatoSeparados = True Then
                Me.txtPrimerNombre.Text = MATRIZ(3)
            End If
            'txtRazonSocial.Text = MATRIZ(0)
            Me.txtDireccion.Text = MATRIZ(4)
            Me.txtCodigoUbigeo.Text = MATRIZ(5)
            
        
        Else
            'Call verificRuc
        End If
        
    ElseIf xWml.Status = 500 Then
        MsgBox "No hay resultados a la consulta", vbInformation, "JNE"
    ElseIf xWml.Status = 12007 Then
        MsgBox "Error en la conexion con Internet", vbInformation, sMensaje
    Else
        MsgBox "Error en el proceso" & xWml.Status, vbInformation, sMensaje
    End If
    
    Set xWml = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
fin:
    MsgBox "Mensaje: " & Error
    Screen.MousePointer = vbDefault

End Sub
'Private Sub verificRuc()
'On Error GoTo fin
'    Dim loRUC As vfpsrucperu.vfpsruc
'    Set loRUC = New vfpsrucperu.vfpsruc
'
'    Dim lcNroRuc As String
'
'        lcNroRuc = txtResumido.Text
'        If loRUC.VFPs_ConsultarRUC(lcNroRuc, False) Then
'            'DEVOLVIO LA CONSULTA CORRECTAMENTE
'            'PROPIEDADES A CONSULTAR LUEGO DE LA CONSULTA DE RUC
'            'loRUC.LCRUC
'            txtDetallado.Text = loRUC.LCRAZONSOCIAL
'            'loRUC.LCTIPOCON
'            'loRUC.c
'            'loRUC.LCTELEFONO
'            'loRUC.LDFECHAINS
'            'loRUC.LCESTADO
'            'loRUC.LCCONDICION
'            txtDireccion.Text = loRUC.LCDIRECCION
'            'loRUC.LDFECHAINICIO
'            'loRUC.LCSISEMICOMP
'            'loRUC.LCACTCOMEXT
'            'loRUC.LCSISCONTA
'        End If
'Exit Sub
'fin:
'    MsgBox "Error al consultar Ruc", vbInformation, sMensaje
'End Sub

Private Sub cmdVerificaDNI_Click()
On Error GoTo fin
    Screen.MousePointer = vbHourglass
    Dim xDat As String
    Dim xRazSoc As String, xEst As String, xCon As String, xDir As String
    Dim xRazSocX As Long, xEstX As Long, xConX As Long, xDirX As Long
    Dim xRazSocY As Long, xEstY As Long, xConY As Long, xDirY As Long
    Dim xWml As New XMLHTTP
    
     Me.txtPrimerNombre.Text = ""
    
    If Not IsNumeric(Trim(txtResumido.Text)) Then
        MsgBox "formato incorrectO, Verificar!!!", vbInformation, sMensaje
        Exit Sub
    End If
    If Len(Trim(txtResumido.Text)) <> 8 Then
        MsgBox "longitud incorrecto, Verificar!!!", vbInformation, sMensaje
        Exit Sub
    End If
    

    
    xWml.Open "POST", "https://cloudservices.infomatica.pe/api/consultadni/" & Trim(txtResumido.Text), False
    xWml.send
    If xWml.Status = 200 Then
        xDat = Trim(xWml.responseText)
        xDat = Replace(xDat, vbNewLine, " ")
        xDat = Replace(Replace(Replace(xDat, "," + Chr(34), vbNewLine), "{", ""), "}", "")
        xDat = Replace(xDat, Chr(34), "")
        
        Dim MATRIZ() As String
        MATRIZ = Split(xDat, vbNewLine)
        
        MATRIZ(0) = Replace(MATRIZ(0), "status:", "")
        If (MATRIZ(0) = 1) Then
            MATRIZ(2) = Replace(MATRIZ(2), "dni:", "")
            MATRIZ(3) = Replace(MATRIZ(3), "apellidos_nombres:", "")
            MATRIZ(4) = Replace(MATRIZ(4), "direccion:", "")
            MATRIZ(5) = Replace(MATRIZ(5), "ubigeo:", "")
            
            'Me.txtUrbanizacion.Text = Calcular("select tdistrito as codigo from tubigeo where tcodigo='" & MATRIZ(5) & "'", Cn)
            Me.txtDetallado.Text = MATRIZ(3)
            
            
            If lDatoSeparados = True Then
                Me.txtPrimerNombre.Text = MATRIZ(3)
            End If
                        
            'txtRazonSocial.Text = MATRIZ(0)
            'Me.txtDireccion.Text = MATRIZ(4)
            'Me.txtCodigoUbigeo.Text = MATRIZ(5)
        
        Else
            'Call verificRuc
        End If
        
    ElseIf xWml.Status = 500 Then
        MsgBox "No hay resultados a la consulta", vbInformation, "JNE"
    ElseIf xWml.Status = 12007 Then
        MsgBox "Error en la conexion con Internet", vbInformation, sMensaje
    Else
        MsgBox "Error en el proceso" & xWml.Status, vbInformation, sMensaje
    End If
    
    Set xWml = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
fin:
    MsgBox "Mensaje: " & Error
    Screen.MousePointer = vbDefault

End Sub

Private Sub Command1_Click()
    Dim xCriterio As String
   Isql = "Select tCodigo as Codigo, tDistrito as Descripcion, tProvincia as Provincia, tDepartamento as Departamento from TUBIGEO order by tCodigo asc"
   
   frmBusca.cboCriterio.Enabled = True
   frmBusca.nPredeterm = 1
   Call ConfGrilla(4, frmBusca.grdGrilla, "Codigo", 2, "Codigo", 1200, 0, 0, "", _
                                          "Distrito", 2, "Descripcion", 1500, 0, 0, "", _
                                          "Provincia", 2, "Provincia", 2500, 0, 0, "", _
                                          "Departamento", 2, "Departamento", 3000, 0, 0, "")
   frmBusca.Show vbModal
   If Not wEnter Then
      Exit Sub
   End If
   txtCodigoUbigeo.Text = sCodigo
End Sub

Private Sub Command2_Click()
'9016
On Error GoTo fin
    Screen.MousePointer = vbHourglass
    Dim xDat As String

    Dim xWml As New XMLHTTP
    
    If Not IsNumeric(Trim(txtResumido.Text)) Then
        MsgBox "formato incorrectO, Verificar!!!", vbInformation, sMensaje
        Exit Sub
    End If
    
    If Len(Trim(txtResumido.Text)) <> 11 Then
        MsgBox "longitud incorrecto, Verificar!!!", vbInformation, sMensaje
        Exit Sub
    End If
    
    xWml.Open "GET", RutaConsultaRucRetenciones & Trim(txtResumido.Text), False
    xWml.send
    
    If xWml.Status = 200 Then
        xDat = Trim(xWml.responseText)
        xDat = Replace(xDat, vbNewLine, " ")
        xDat = Replace(Replace(Replace(xDat, "," + Chr(34), vbNewLine), "{", ""), "}", "")
        xDat = Replace(xDat, Chr(34), "")
        
        Dim MATRIZ() As String
        MATRIZ = Split(xDat, vbNewLine)
        
         
        
        MATRIZ(0) = Replace(MATRIZ(3), "status:", "")
       If (MATRIZ(0) <> "Ruc:null") Then
            chkAgenteRetencion.Value = 1
        Else
            chkAgenteRetencion.Value = 0
        End If
        
    ElseIf xWml.Status = 500 Then
        MsgBox "No hay resultados a la consulta", vbInformation, "JNE"
    ElseIf xWml.Status = 12007 Then
        MsgBox "Error en la conexion con Internet", vbInformation, sMensaje
    Else
        MsgBox "Error en el proceso" & xWml.Status, vbInformation, sMensaje
    End If
    
    Set xWml = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
fin:
    MsgBox "Mensaje: " & Error
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()

   ' cambios validar DNI
   Isql = "select lValidaDNI from TPARAMETRO"
   Set RsParametro = Lib.OpenRecordset(Isql, Cn)
   '--------------------------------------------
   
    Screen.MousePointer = vbHourglass
    Centrar Me
    LlenaCombos
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Clientes / Transportistas "
    fraDetalle.Caption = Me.Caption
    
    If lSAP Then
        lblTipoCliente.Visible = True
        cboTipoCliente.Visible = True
    Else
        lblTipoCliente.Visible = False
        cboTipoCliente.Visible = False
    End If
    
    If pais = "000" Then
        cmdVerifica.Visible = True
        cmdVerificaDNI.Visible = True
    ElseIf pais = "003" Then
        cboTipoResponsable.Visible = True
        txtCodigoUbigeo.Visible = False
        Me.Command1.Visible = False
        cmdVerifica.Visible = False
    Else
        cmdVerifica.Visible = False
        cmdVerificaDNI.Visible = False
    End If
    
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.Value = 1
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       'frmClienteFactura.RsCabecera.Find ("Codigo = '" & frmClienteFactura.RsCabecera!Codigo & "'")
       ActivarBotones (True)
       Asignar
    End If
    
    If Not lClub Then
        Label4.Visible = False
        txtEnlace.Visible = False
    End If
    
    cmdTexto.Caption = "Registro " & frmClienteFactura.RsCabecera.AbsolutePosition & " de " & frmClienteFactura.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Sub LlenaCombos()
    With cboTipoIdentidad
         If lagregar Then
            Isql = "Select * from vTipoIdentidad where lactivo=1 order by Descripcion"
         Else
            Isql = "Select * from vTipoIdentidad order by Descripcion"
         End If
         
         Set RsTipoIdentidad = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoIdentidad
             .DataField = "tResumido"
             .ListField = "tResumido"
             .BoundColumn = "Codigo"
    End With
    
    With cboTipoResponsable
         If lagregar Then
            Isql = "Select * from vTipoResponsable where lactivo=1 order by codigo"
         Else
            Isql = "Select * from vTipoResponsable order by codigo"
         End If
         
         Set RsTipoResponsable = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoResponsable
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
    With cboTipoCliente
         If lagregar Then
            Isql = "Select * from vTipoGrupoCliente where lactivo=1 order by Descripcion"
         Else
            Isql = "Select * from vTipoGrupoCliente order by Descripcion"
         End If
         
         Set RsTipoCliente = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoCliente
             .DataField = "tResumido"
             .ListField = "tResumido"
             .BoundColumn = "Codigo"
    End With
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
        Set RsTipoIdentidad = Nothing
        Set RsTipoCliente = Nothing

    Set frmClienteFacturaDetalle = Nothing
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




Private Sub txtDetallado_LostFocus()
  ' Call ValidaStr(txtDetallado)
End Sub

Private Sub txtDireccion_LostFocus()
   Call ValidaStr(txtDireccion)
End Sub

Private Sub txtPrimerApellido_Change()
On Error GoTo fin
    If lDatoSeparados Then
        txtDetallado.Text = Trim(Me.txtPrimerNombre) & " " & Trim(Me.txtSegundoNombre) & " " & Trim(Me.txtPrimerApellido) & " " & Trim(Me.txtSegundoApellido)
    End If
Exit Sub
fin:
End Sub

Private Sub txtPrimerApellido_KeyPress(KeyAscii As Integer)
    If NadaSimbolos(KeyAscii) = False Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txtPrimerNombre_Change()
On Error GoTo fin
     If lDatoSeparados Then
        txtDetallado.Text = Trim(Me.txtPrimerNombre) & " " & Trim(Me.txtSegundoNombre) & " " & Trim(Me.txtPrimerApellido) & " " & Trim(Me.txtSegundoApellido)
    End If
    Exit Sub
fin:
End Sub

Private Sub txtPrimerNombre_KeyPress(KeyAscii As Integer)
    If NadaSimbolos(KeyAscii) = False Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txtResumido_LostFocus()
    Call ValidaStr(txtResumido)
    If cboTipoIdentidad.BoundText = "02" Then
    'Consitencia RUC
            If lLongitud And Len(Trim(txtResumido.Text)) <> nLongitud Then
               MsgBox "La longitud del Id. Tributario debe ser " & nLongitud, vbCritical, sMensaje
               Exit Sub
            ElseIf Not lLongitud And Len(Trim(txtResumido.Text)) < nLongitud Then
               MsgBox "La longitud del Identificador debe ser mayor igual a " & nLongitud, vbCritical, sMensaje
               Exit Sub
            End If
    End If
End Sub

'cambio de validar DNI
Private Sub txtResumido_GotFocus()
If cboTipoIdentidad = "" Then
    MsgBox "Debe colocar un identificador"
    foco
    Exit Sub
End If
End Sub

'cambio de validar DNI
Private Function foco()
    txtDetallado.SetFocus
End Function


'cambio de validar DNI
Private Sub txtResumido_KeyPress(KeyAscii As Integer)
    If RsParametro!lValidaDNI = True Then
        If cboTipoIdentidad.SelectedItem = 2 Then
            If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{tab}"
            ElseIf KeyAscii <> 8 Then
            If Not IsNumeric(Chr(KeyAscii)) Then
            Beep
            KeyAscii = 0
            End If
            End If
        End If
    End If
    
    
    If NadaSimbolos(KeyAscii) = False Then
        Beep
        KeyAscii = 0
    End If
    
    
    
End Sub

Private Sub txtSegundoApellido_Change()
On Error GoTo fin
       If lDatoSeparados Then
        txtDetallado.Text = Trim(Me.txtPrimerNombre) & " " & Trim(Me.txtSegundoNombre) & " " & Trim(Me.txtPrimerApellido) & " " & Trim(Me.txtSegundoApellido)
    End If
    Exit Sub
fin:
End Sub

Private Sub txtSegundoApellido_KeyPress(KeyAscii As Integer)
    If NadaSimbolos(KeyAscii) = False Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txtSegundoNombre_Change()
On Error GoTo fin
      If lDatoSeparados Then
        txtDetallado.Text = Trim(Me.txtPrimerNombre) & " " & Trim(Me.txtSegundoNombre) & " " & Trim(Me.txtPrimerApellido) & " " & Trim(Me.txtSegundoApellido)
    End If
    Exit Sub
fin:
End Sub

Private Sub txtSegundoNombre_KeyPress(KeyAscii As Integer)
    If NadaSimbolos(KeyAscii) = False Then
        Beep
        KeyAscii = 0
    End If
End Sub
' Función IsLetter
Private Function IsLetter(ByVal char As String) As Boolean
    IsLetter = (Asc(char) >= 65 And Asc(char) <= 90) Or (Asc(char) >= 97 And Asc(char) <= 122)
End Function
Private Function IncrementarLetras(ByVal letras As String) As String
    Dim i As Integer
    Dim arr() As Byte
    Dim resultado As String
    
    If letras = "" Then
        resultado = "A"
    Else
        arr = StrConv(letras, vbFromUnicode)
        
        For i = UBound(arr) To LBound(arr) Step -1
            If arr(i) = Asc("Z") Then
                arr(i) = Asc("A")
                If i = LBound(arr) Then
                    resultado = "A" & StrConv(arr, vbUnicode)
                End If
            Else
                arr(i) = arr(i) + 1
                resultado = StrConv(arr, vbUnicode)
                Exit For
            End If
        Next i
    End If
    
    IncrementarLetras = resultado
End Function
