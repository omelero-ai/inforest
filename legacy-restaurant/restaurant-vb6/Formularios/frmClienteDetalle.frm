VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmClienteDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5010
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   11610
   Icon            =   "frmClienteDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   11610
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
      Height          =   4185
      Left            =   2580
      TabIndex        =   27
      Top             =   30
      Width           =   8955
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
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   2
         Text            =   " "
         Top             =   795
         Width           =   7095
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
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   3
         Text            =   " "
         Top             =   1155
         Width           =   7095
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
         Left            =   1665
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
         Left            =   990
         TabIndex        =   13
         Top             =   3735
         Width           =   885
      End
      Begin VB.CheckBox ChkEmpresa 
         Alignment       =   1  'Right Justify
         Caption         =   "Empresa :"
         Height          =   195
         Left            =   3150
         TabIndex        =   1
         Top             =   435
         Width           =   1110
      End
      Begin VB.TextBox txtIdentidad 
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
         Left            =   6030
         MaxLength       =   15
         TabIndex        =   5
         Text            =   " "
         Top             =   1530
         Width           =   2730
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
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   6
         Text            =   " "
         Top             =   1875
         Width           =   7095
      End
      Begin VB.TextBox txtTelefono1 
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
         Left            =   1665
         MaxLength       =   15
         TabIndex        =   7
         Text            =   " "
         Top             =   2235
         Width           =   2730
      End
      Begin VB.TextBox txtTelefono2 
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
         Left            =   6030
         MaxLength       =   15
         TabIndex        =   8
         Text            =   " "
         Top             =   2235
         Width           =   2730
      End
      Begin VB.TextBox txtEmail 
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
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   9
         Text            =   " "
         Top             =   2595
         Width           =   7095
      End
      Begin VB.TextBox txtLinea 
         Alignment       =   1  'Right Justify
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
         Left            =   1665
         MaxLength       =   20
         TabIndex        =   10
         Text            =   " "
         Top             =   2955
         Width           =   2730
      End
      Begin MSDataListLib.DataCombo cboTipoIdentidad 
         Height          =   315
         Left            =   1665
         TabIndex        =   4
         Top             =   1515
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
      Begin MSDataListLib.DataCombo cboTipoCtaCte 
         Height          =   315
         Left            =   1665
         TabIndex        =   11
         Top             =   3330
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
      Begin MSDataListLib.DataCombo cboSubTipoCtaCte 
         Height          =   315
         Left            =   6030
         TabIndex        =   12
         Top             =   3330
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
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "SubTipo Cta.Cte. :"
         Height          =   195
         Index           =   11
         Left            =   4635
         TabIndex        =   40
         Top             =   3390
         Width           =   1305
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono (2) :"
         Height          =   195
         Index           =   10
         Left            =   4995
         TabIndex        =   39
         Top             =   2280
         Width           =   945
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Número Identidad :"
         Height          =   195
         Index           =   9
         Left            =   4590
         TabIndex        =   38
         Top             =   1575
         Width           =   1350
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cta.Cte. :"
         Height          =   195
         Index           =   8
         Left            =   555
         TabIndex        =   37
         Top             =   3375
         Width           =   1020
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Razón Social :"
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   35
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Razón Comercial :"
         Height          =   195
         Index           =   2
         Left            =   285
         TabIndex        =   34
         Top             =   1200
         Width           =   1290
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Identidad :"
         Height          =   195
         Index           =   3
         Left            =   465
         TabIndex        =   33
         Top             =   1560
         Width           =   1110
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   990
         TabIndex        =   32
         Top             =   435
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono (1) :"
         Height          =   195
         Index           =   5
         Left            =   630
         TabIndex        =   31
         Top             =   2280
         Width           =   945
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Dirección :"
         Height          =   195
         Index           =   4
         Left            =   810
         TabIndex        =   30
         Top             =   1920
         Width           =   765
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail :"
         Height          =   195
         Index           =   6
         Left            =   1050
         TabIndex        =   29
         Top             =   2640
         Width           =   525
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Linea de Crédito :"
         Height          =   195
         Index           =   7
         Left            =   330
         TabIndex        =   28
         Top             =   3000
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   11550
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4260
      Width           =   11610
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Tiendas"
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
         Index           =   4
         Left            =   7995
         Picture         =   "frmClienteDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Left            =   10335
         Picture         =   "frmClienteDetalle.frx":053C
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Left            =   9165
         Picture         =   "frmClienteDetalle.frx":062E
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Left            =   5655
         Picture         =   "frmClienteDetalle.frx":0730
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Left            =   6825
         Picture         =   "frmClienteDetalle.frx":0C62
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   5490
         TabIndex        =   20
         Top             =   60
         Width           =   5550
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmClienteDetalle.frx":1194
            Style           =   1  'Graphical
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmClienteDetalle.frx":16D6
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmClienteDetalle.frx":1C18
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5010
            Picture         =   "frmClienteDetalle.frx":215A
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4530
            Picture         =   "frmClienteDetalle.frx":269C
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4005
            Picture         =   "frmClienteDetalle.frx":2BDE
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   0
            Width           =   525
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
            Left            =   1440
            TabIndex        =   36
            Top             =   150
            Width           =   2520
         End
      End
   End
   Begin VB.Image Image 
      Height          =   4140
      Left            =   45
      Picture         =   "frmClienteDetalle.frx":3120
      Stretch         =   -1  'True
      Top             =   60
      Width           =   2490
   End
End
Attribute VB_Name = "frmClienteDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTipoIdentidad As Recordset
Dim nLinea As Double
Dim nCodigo As String
Dim sRUC As String
Dim RsTipoCtaCte As Recordset
Dim RsSubTipoCtaCte As Recordset

Sub LlenaCombos()
    
    With cboTipoIdentidad
         Isql = "Select * from vTipoIdentidad order by Descripcion"
         Set RsTipoIdentidad = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoIdentidad
             .DataField = "tResumido"
             .ListField = "tResumido"
             .BoundColumn = "Codigo"
    End With
    
    With cboTipoCtaCte
         Isql = "Select * from vTipoCtaCte where lActivo=1 order by Descripcion"
         Set RsTipoCtaCte = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsTipoCtaCte
             .DataField = "tResumido"
             .ListField = "tResumido"
             .BoundColumn = "Codigo"
    End With
    
    With cboSubTipoCtaCte
         Isql = "Select * from vSubTipoCtaCte where lActivo=1 order by Descripcion"
         Set RsSubTipoCtaCte = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubTipoCtaCte
             .DataField = "tResumido"
             .ListField = "tResumido"
             .BoundColumn = "Codigo"
    End With
    
End Sub

Sub Asignar()
    With frmCliente.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!tApeCom), "", !tApeCom)
        txtResumido = IIf(IsNull(!tNomSoc), "", !tNomSoc)
        txtIdentidad = IIf(IsNull(!tIdentidad), "", !tIdentidad)
        sRUC = IIf(IsNull(!tIdentidad), "", !tIdentidad)
        txtDireccion = IIf(IsNull(!tDireccion), "", !tDireccion)
        txtTelefono1 = IIf(IsNull(!ttelefono1), "", !ttelefono1)
        txtTelefono2 = IIf(IsNull(!ttelefono2), "", !ttelefono2)
        txtEmail = IIf(IsNull(!temail), "", !temail)
        nLinea = IIf(IsNull(!nLinea), 0, !nLinea)
        txtLinea = Format(nLinea, "###,##0.00")
    
        'Data Combo
        cboTipoIdentidad.BoundText = IIf(IsNull(!tTipoIdentidad), "", Trim(!tTipoIdentidad))
        cboTipoCtaCte.BoundText = IIf(IsNull(!tTipoCtaCte), "", Trim(!tTipoCtaCte))
        cboSubTipoCtaCte.BoundText = IIf(IsNull(!tSubTipoCtaCte), "", Trim(!tSubTipoCtaCte))
        
        'Check Box
        ChkEmpresa = IIf(!lEmpresa = True, 1, 0)
        If ChkEmpresa.value = 1 Then
           Label(1).Caption = "Razón Comercial :"
           Label(2).Caption = "Razón Social :"
        Else
           Label(1).Caption = "Apellidos :"
           Label(2).Caption = "Nombres :"
        End If
        chkActivo = IIf(!lActivo = True, 1, 0)
    End With
End Sub

Private Sub cboTipoCtaCte_Click(Area As Integer)
    cboSubTipoCtaCte.BoundText = ""
    With cboSubTipoCtaCte
         Isql = "Select * from vSubtipoctacte where tTipoCtaCte = '" & cboTipoCtaCte.BoundText & "' order by Descripcion "
         Set RsSubTipoCtaCte = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSubTipoCtaCte
    End With
End Sub

Private Sub ChkEmpresa_Click()
    If ChkEmpresa.value = 1 Then
       Label(1).Caption = "Razón Comercial :"
       Label(2).Caption = "Razón Social :"
    Else
       Label(1).Caption = "Apellidos :"
       Label(2).Caption = "Nombres :"
    End If
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmCliente.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmCliente.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmCliente.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmCliente.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmCliente.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmCliente.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmCliente.RsCabecera.RecordCount = 0, 0, frmCliente.RsCabecera.AbsolutePosition) & " de " & frmCliente.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               chkActivo.value = 1
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
                    
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               
               cmdTexto.Caption = "Registro " & IIf(frmCliente.RsCabecera.RecordCount = 0, 0, frmCliente.RsCabecera.AbsolutePosition) & " de " & frmCliente.RsCabecera.RecordCount
               'Chequea Datos
               If ChkEmpresa = 1 Then
                  If txtDetallado.Text = "" Then MsgBox "Ingrese la Razón Comercial", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
                  If txtResumido.Text = "" Then MsgBox "Ingrese la Razón Social", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
               Else
                  If txtDetallado.Text = "" Then MsgBox "Ingrese los Apellidos", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
                  If txtResumido.Text = "" Then MsgBox "Ingrese los Nombres", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
               End If
               
               'If ChkEmpresa = 1 Then
                If Calcular("select nValor as Codigo from vTipoIdentidad where Codigo='" & cboTipoIdentidad.BoundText & "'", Cn) = 1 Then
                  If lLongitud And Len(Trim(txtIdentidad.Text)) <> nLongitud Then
                     MsgBox "La longitud del Identificador debe ser " & nLongitud, vbCritical, sMensaje
                     Exit Sub
                  ElseIf Not lLongitud And Len(Trim(txtIdentidad.Text)) < nLongitud Then
                     MsgBox "La longitud del Identificador debe ser Mayor a " & nLongitud, vbCritical, sMensaje
                     Exit Sub
                  End If
               End If
               
               If Sw Then
                  If Calcular("select nValor as Codigo from vTipoIdentidad where Codigo='" & cboTipoIdentidad.BoundText & "'", Cn) = 1 Then
                     If Not Calcular("select tIdentidad as Codigo from TCOMPANIA where tIdentidad ='" & Trim(txtIdentidad.Text) & "'", Cn) = "0" Then
                        MsgBox "Identificador Tributario existente", vbExclamation, sMensaje
                        Exit Sub
                     End If
                  End If
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigoCliente) as Codigo from TCOMPANIA", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "00001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 5)
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
                  Isql = "insert into TCOMPANIA( " & _
                         "tCodigoCliente, lEmpresa, tApeCom, tNomSoc, tTipoIdentidad, tIdentidad, " & _
                         "tDireccion, tTelefono1, tTelefono2, tEmail, nLinea, nConsumo, lActivo, tUsuario, tTipoCtaCte, tSubTipoCtaCte, fRegistro) " & _
                         "values ('" & txtCodigo.Text & "', " & _
                                       ChkEmpresa.value & ", " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                " '" & cboTipoIdentidad.BoundText & "', " & _
                                " '" & txtIdentidad.Text & "', " & _
                                " '" & txtDireccion.Text & "', " & _
                                " '" & txtTelefono1.Text & "', " & _
                                " '" & txtTelefono2.Text & "', " & _
                                " '" & txtEmail.Text & "', " & _
                                "  " & nLinea & ", 0, " & _
                                       chkActivo.value & ", " & _
                                  " '" & sUsuario & "', " & _
                                  " '" & cboTipoCtaCte.BoundText & "', " & _
                                  " '" & cboSubTipoCtaCte.BoundText & "', " & _
                                  " getdate() )"
                  Cn.Execute Isql
                  frmCliente.RsCabecera.Sort = "Codigo ASC"
                  frmCliente.RsCabecera.Requery
                  frmCliente.RsCabecera.MoveLast
                                                      
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmCliente.RsCabecera.RecordCount = 0, 0, frmCliente.RsCabecera.AbsolutePosition) & " de " & frmCliente.RsCabecera.RecordCount
               Else
                  'Cambiar el SQL
                  If Calcular("select nValor as Codigo from vTipoIdentidad where Codigo='" & cboTipoIdentidad.BoundText & "'", Cn) = 1 Then
                     If Calcular("select tIdentidad as Codigo from TCOMPANIA where tIdentidad ='" & Trim(txtIdentidad.Text) & "' and tCodigoCliente <> '" & txtCodigo.Text & "'", Cn) > 0 Then
                        MsgBox "Identificador Tributario existente", vbExclamation, sMensaje
                        Exit Sub
                     End If
                  End If
                  
                  Isql = "update TCOMPANIA set " & _
                         "tApeCom ='" & txtDetallado.Text & "', " & _
                         "tNomSoc ='" & txtResumido.Text & "', " & _
                         "tTipoIdentidad ='" & cboTipoIdentidad.BoundText & "', " & _
                         "tTipoCtaCte ='" & cboTipoCtaCte.BoundText & "', " & _
                         "tsubTipoCtaCte ='" & cboSubTipoCtaCte.BoundText & "', " & _
                         "tIdentidad ='" & txtIdentidad.Text & "', " & _
                         "tDireccion ='" & txtDireccion.Text & "', " & _
                         "tTelefono1 ='" & txtTelefono1.Text & "', " & _
                         "tTelefono2 ='" & txtTelefono2.Text & "', " & _
                         "tEmail ='" & txtEmail.Text & "', " & _
                         "nLinea =" & nLinea & ", " & _
                         "lEmpresa =" & ChkEmpresa.value & ", " & _
                         "lActivo =" & chkActivo.value & _
                         " where tCodigoCliente = '" & txtCodigo & "'"
                   
                   Cn.Execute Isql
                   nPos = frmCliente.RsCabecera.Bookmark
                   frmCliente.RsCabecera.Requery
                   If frmCliente.RsCabecera.RecordCount = 0 Then
                      frmCliente.RsCabecera.Filter = adFilterNone
                   End If
                   frmCliente.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
                   
          Case Is = 2 ' Eliminar
               If frmCliente.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Cliente " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               If Calcular("select count(tdocumento) as Codigo from MDOCUMENTO where tClientePago='" & txtCodigo & "'", Cn) > 0 Then
                  MsgBox "Cliente con movimientos, no se puede eliminar", vbExclamation, sMensaje
                  Exit Sub
               Else
                  Cn.Execute "delete from TCOMPANIA where tCodigoCliente = '" & txtCodigo & "'"
                  frmCliente.RsCabecera.Requery
               End If
               If frmCliente.RsCabecera.RecordCount <> 0 Then
                  frmCliente.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmCliente.RsCabecera.RecordCount = 0, 0, frmCliente.RsCabecera.AbsolutePosition) & " de " & frmCliente.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If
          
          Case Is = 3 ' Salir
               Unload Me
               
          Case Is = 4 ' Tiendas
               If txtCodigo.Text = "" Then
                  Exit Sub
               End If
               sCodigo = txtCodigo.Text
               sDescrip = txtDetallado.Text
               sTemp = txtResumido.Text
               frmTienda.Show vbModal
   End Select

End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Clientes "
    fraDetalle.Caption = Me.Caption
    
    'Llena todos los Combos
    LlenaCombos
        
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.value = 1
       ChkEmpresa.value = 1
    Else
       'Cambiar el Nombre del formulario Cabecera
       'frmCliente.RsCabecera.Find ("Codigo = '" & frmCliente.RsCabecera!Codigo & "'")
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & IIf(frmCliente.RsCabecera.RecordCount = 0, 0, frmCliente.RsCabecera.AbsolutePosition) & " de " & frmCliente.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmClienteDetalle = Nothing
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

Private Sub txtLinea_LostFocus()
    nLinea = Val(txtLinea.Text)
    txtLinea.Text = Format(nLinea, "###,##0.00")
End Sub
