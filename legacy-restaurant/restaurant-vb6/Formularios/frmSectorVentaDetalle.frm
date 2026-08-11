VERSION 5.00
Begin VB.Form frmSectorVentaDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3315
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9510
   Icon            =   "frmSectorVentaDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   9510
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
      Height          =   2535
      Left            =   1800
      TabIndex        =   18
      Top             =   0
      Width           =   7665
      Begin VB.TextBox txtCuentaContable 
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
         Left            =   2100
         MaxLength       =   20
         TabIndex        =   7
         Text            =   " "
         Top             =   1710
         Width           =   3555
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   1250
         TabIndex        =   8
         Top             =   2160
         Width           =   1050
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
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   1170
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
         Left            =   2100
         MaxLength       =   20
         TabIndex        =   6
         Text            =   " "
         Top             =   1230
         Width           =   3555
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
         Left            =   2100
         MaxLength       =   35
         TabIndex        =   5
         Text            =   " "
         Top             =   765
         Width           =   5490
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cuenta Contable :"
         Height          =   195
         Left            =   570
         TabIndex        =   22
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   1260
         TabIndex        =   21
         Top             =   465
         Width           =   585
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Detallada :"
         Height          =   195
         Left            =   195
         TabIndex        =   20
         Top             =   810
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Resumida :"
         Height          =   195
         Left            =   165
         TabIndex        =   19
         Top             =   1320
         Width           =   1680
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9450
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2565
      Width           =   9510
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
         Picture         =   "frmSectorVentaDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "frmSectorVentaDetalle.frx":0534
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "frmSectorVentaDetalle.frx":0636
         Style           =   1  'Graphical
         TabIndex        =   0
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
         Picture         =   "frmSectorVentaDetalle.frx":0B68
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4590
         TabIndex        =   16
         Top             =   60
         Width           =   4650
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmSectorVentaDetalle.frx":109A
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmSectorVentaDetalle.frx":15DC
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmSectorVentaDetalle.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4110
            Picture         =   "frmSectorVentaDetalle.frx":2060
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3630
            Picture         =   "frmSectorVentaDetalle.frx":25A2
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3150
            Picture         =   "frmSectorVentaDetalle.frx":2AE4
            Style           =   1  'Graphical
            TabIndex        =   12
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
            Left            =   1500
            TabIndex        =   17
            Top             =   180
            Width           =   1665
         End
      End
   End
   Begin VB.Image Image 
      Height          =   2490
      Left            =   15
      Picture         =   "frmSectorVentaDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   45
      Width           =   1755
   End
End
Attribute VB_Name = "frmSectorVentaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSector As Recordset

Sub Asignar()
    With frmSectorVenta.RsCabecera
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Detallado), "", !Detallado)
        txtResumido = IIf(IsNull(!Resumido), "", !Resumido)
        txtCuentaContable = IIf(IsNull(!CuentaContable), "", !CuentaContable)
        chkActivo = IIf(!Activo = True, 1, 0)
    End With
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmSectorVenta.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmSectorVenta.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmSectorVenta.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmSectorVenta.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmSectorVenta.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmSectorVenta.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmSectorVenta.RsCabecera.RecordCount = 0, 0, frmSectorVenta.RsCabecera.AbsolutePosition) & " de " & frmSectorVenta.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               chkActivo.value = 1
               txtDetallado.SetFocus
                    
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
     
               
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTabla where ttabla='SECTORVENTA' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "01"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 2)
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
                  Isql = " INSERT INTO TTABLA (TTABLA, tCodigo,tDetallado,tResumido, tValor, lActivo,lReplica) " & _
                         " values ('SECTORVENTA', '" & txtCodigo.Text & "', " & _
                         " '" & UCase(txtDetallado.Text) & "', " & _
                         " '" & UCase(txtResumido.Text) & "', " & _
                         " '" & txtCuentaContable.Text & "', " & _
                         chkActivo.value & ", 1 ) "
                                       

                  Cn.Execute Isql
               '   Cn.Execute Isql
                  frmSectorVenta.RsCabecera.Sort = "Codigo ASC"
                  frmSectorVenta.RsCabecera.Requery
                  frmSectorVenta.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmSectorVenta.RsCabecera.RecordCount = 0, 0, frmSectorVenta.RsCabecera.AbsolutePosition) & " de " & frmSectorVenta.RsCabecera.RecordCount
               Else
                  'Cambiar el SQL
       
                        Isql = "update TTABLA set " & _
                         "tDetallado ='" & UCase(txtDetallado.Text) & "', " & _
                         "tResumido ='" & UCase(txtResumido.Text) & "', " & _
                         "tValor ='" & txtCuentaContable.Text & "', " & _
                         "lActivo =" & chkActivo.value & ", lReplica=1 " & _
                         " where TTABLA='SECTORVENTA' AND tCodigo = '" & txtCodigo & "'"

                  Cn.Execute Isql
   
                  nPos = frmSectorVenta.RsCabecera.Bookmark
                  frmSectorVenta.RsCabecera.Requery
                  If frmSectorVenta.RsCabecera.RecordCount = 0 Then
                     frmSectorVenta.RsCabecera.Filter = adFilterNone
                  End If
                  frmSectorVenta.RsCabecera.Bookmark = nPos
                  Screen.MousePointer = vbDefault
                  MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          Case Is = 2 ' Eliminar
               If frmSectorVenta.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Sector Venta " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               Dim calculo As Long
               calculo = Calcular("select COUNT(*) AS codigo from TCAJA where tSectorVenta='" & txtCodigo & "'", Cn)
               If calculo = 0 Then
                    'Cambia el Delete
                    Cn.Execute "delete from ttabla where ttabla='SECTORVENTA' AND  tcodigo='" & txtCodigo & "'"
                    frmSectorVenta.RsCabecera.Requery
                    If frmSectorVenta.RsCabecera.RecordCount <> 0 Then
                       frmSectorVenta.RsCabecera.MoveLast
                       Asignar
                       cmdTexto.Caption = "Registro " & IIf(frmSectorVenta.RsCabecera.RecordCount = 0, 0, frmSectorVenta.RsCabecera.AbsolutePosition) & " de " & frmSectorVenta.RsCabecera.RecordCount
                    Else
                       ActivarBotones False
                       Blanquear Me
                       Sw = True
                    End If
                Else
                    MsgBox "No se puede eliminar el registro ", vbCritical + vbInformation, sMensaje
                End If
          Case Is = 3 ' Salir
               Unload Me
   End Select

End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
    'Ingrese el SubTitulo
    Me.Caption = " Agrupacion de Puntos de Ventas "
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
    
    cmdTexto.Caption = "Registro " & IIf(frmSectorVenta.RsCabecera.RecordCount = 0, 0, frmSectorVenta.RsCabecera.AbsolutePosition) & " de " & frmSectorVenta.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmSectorVentaDetalle = Nothing
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


