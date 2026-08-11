VERSION 5.00
Begin VB.Form frmCentroCostoDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2475
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   10230
   Icon            =   "frmCentroCostoDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   10230
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
      Height          =   1700
      Left            =   2160
      TabIndex        =   14
      Top             =   0
      Width           =   8050
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   360
         Width           =   1170
      End
      Begin VB.TextBox txtCentroCosto 
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
         Left            =   2760
         MaxLength       =   120
         TabIndex        =   1
         Text            =   " "
         Top             =   795
         Width           =   4980
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   2040
         TabIndex        =   2
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   2025
         TabIndex        =   18
         Top             =   405
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo :"
         Height          =   195
         Index           =   1
         Left            =   1380
         TabIndex        =   17
         Top             =   840
         Width           =   1230
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   10170
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1725
      Width           =   10230
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
         Left            =   8850
         Picture         =   "frmCentroCostoDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   7680
         Picture         =   "frmCentroCostoDetalle.frx":0534
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   5340
         Picture         =   "frmCentroCostoDetalle.frx":0636
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
         Left            =   6510
         Picture         =   "frmCentroCostoDetalle.frx":0B68
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   5190
         TabIndex        =   7
         Top             =   60
         Width           =   5250
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmCentroCostoDetalle.frx":109A
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmCentroCostoDetalle.frx":15DC
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmCentroCostoDetalle.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4710
            Picture         =   "frmCentroCostoDetalle.frx":2060
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4230
            Picture         =   "frmCentroCostoDetalle.frx":25A2
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3750
            Picture         =   "frmCentroCostoDetalle.frx":2AE4
            Style           =   1  'Graphical
            TabIndex        =   8
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
            Left            =   1440
            TabIndex        =   15
            Top             =   180
            Width           =   2265
         End
      End
   End
   Begin VB.Image Image 
      Height          =   1700
      Left            =   45
      Picture         =   "frmCentroCostoDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2100
   End
End
Attribute VB_Name = "frmCentroCostoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Asignar()
    With frmCentroCosto.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!tCodigo), "", !tCodigo)
        txtCentroCosto = IIf(IsNull(!CentroCosto), "", !CentroCosto)
        chkActivo = IIf(!lActivo = True, 1, 0)
        
    End With
End Sub

Private Sub chkCliente_Click()
   If chkCliente.Value = False Then
      txtMonto.Enabled = False
      txtMonto.Text = "0"
   Else
      txtMonto.Enabled = True
   End If
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmCentroCosto.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmCentroCosto.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmCentroCosto.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmCentroCosto.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmCentroCosto.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmCentroCosto.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmCentroCosto.RsCabecera.RecordCount = 0, 0, frmCentroCosto.RsCabecera.AbsolutePosition) & " de " & frmCentroCosto.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               chkActivo.Value = 1
               'Cambia el Nombre del Primer Text
               txtCentroCosto.SetFocus
                    
          Case Is = 1 ' Grabar
               Dim nCorrela As String
                    
               'Chequea Datos
               If txtCentroCosto.Text = "" Then MsgBox "Ingrese el Código de Autorización", vbExclamation, sMensaje: txtCentroCosto.SetFocus: Exit Sub

               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(TCodigo) as Codigo from TCENTROCOSTO", Cn)
                  
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "000001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 6)
                  End If
                  
                  Sw = False
                   
                  'Cambiar el SQL
'                  Isql = "insert into TCODIGOAUTORIZACION( " & _
'                         "tCodigo, tCodigoAutorizacion, fInicial, fFinal, fRegistro, tusuario, lActivo) " & _
'                         "values  ('" & txtCodigo.Text & "', " & _
'                                " '" & txtCodigoAutorizacion.Text & "', " & _
'                                " '" & Format(dtpFechaIni.Value, "yyyy/MM/dd") & "', " & _
'                                " '" & Format(dtpFechaFin.Value, "yyyy/MM/dd") & "', " & _
'                                " getdate(), " & _
'                                " '" & sUsuario & "', " & _
'                                chkActivo.Value & ") "
'                  Cn.Execute Isql

                  Cn.Execute "exec spIns_CENTROCOSTO ''," & 1 & " , '" & txtCentroCosto.Text & "'," & chkActivo.Value
                  
                  frmCentroCosto.RsCabecera.Sort = "TCodigo ASC"
                  frmCentroCosto.RsCabecera.Requery
                  frmCentroCosto.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmCentroCosto.RsCabecera.RecordCount = 0, 0, frmCentroCosto.RsCabecera.AbsolutePosition) & " de " & frmCentroCosto.RsCabecera.RecordCount
               Else
                  'Cambiar el SQL
'                  Isql = "update TCODIGOAUTORIZACION set " & _
'                                " tCodigoAutorizacion = '" & txtCodigoAutorizacion.Text & "', " & _
'                                " fInicial = '" & Format(dtpFechaIni.Value, "yyyy/MM/dd") & "', " & _
'                                " fFinal = '" & Format(dtpFechaFin.Value, "yyyy/MM/dd") & "', " & _
'                                " fRegistro = getdate(), " & _
'                                " tusuario = '" & sUsuario & "', " & _
'                                " lActivo =" & chkActivo.Value & _
'                         " where tCodigo = '" & txtCodigo & "'"
'                   Cn.Execute Isql
                   Cn.Execute "exec spIns_CENTROCOSTO '" & txtCodigo.Text & "'," & 2 & " , '" & txtCentroCosto.Text & "'," & chkActivo.Value
                   nPos = frmCentroCosto.RsCabecera.Bookmark
                   frmCentroCosto.RsCabecera.Requery
                   If frmCentroCosto.RsCabecera.RecordCount = 0 Then
                      frmCentroCosto.RsCabecera.Filter = adFilterNone
                   End If
                   frmCentroCosto.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          
          Case Is = 2 ' Eliminar
               If frmCentroCosto.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Centro de Costo " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "exec spIns_CENTROCOSTO '" & txtCodigo.Text & "'," & 3 & ",'',''"
               frmCentroCosto.RsCabecera.Requery
               If frmCentroCosto.RsCabecera.RecordCount <> 0 Then
                  frmCentroCosto.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmCentroCosto.RsCabecera.RecordCount = 0, 0, frmCentroCosto.RsCabecera.AbsolutePosition) & " de " & frmCentroCosto.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
                  
                  
               End If

          
          Case Is = 3 ' Salir
               Unload Me
   End Select

End Sub

Private Sub Form_Load()
    Centrar Me
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Centro de Costo"
'    fraDetalle.Caption = Me.Caption
    
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.Value = 1
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & IIf(frmCentroCosto.RsCabecera.RecordCount = 0, 0, frmCentroCosto.RsCabecera.AbsolutePosition) & " de " & frmCentroCosto.RsCabecera.RecordCount

'    'PARA LA GRILLA
'    Dim RsDetalleAut As Recordset
'    Isql = "SELECT TipoDocumento,Serie,FolioInicial, FolioFinal, Caja FROM vAutorizacionDetalle"
'    Set RsDetalleAut = Lib.OpenRecordset(Isql, Cn)
'
'    'Configuración de la Grilla
'    Call ConfGrilla(5, grdGrilla, "TipoDocumento", 2, "TipoDocumento", 1600, 2, 0, "", _
'                                  "Serie", 2, "Serie", 1400, 2, 0, "", _
'                                  "FolioInicial", 2, "FolioInicial", 1500, 2, 0, "", _
'                                  "FolioFinal", 2, "FolioFinal", 1500, 2, 0, "", _
'                                  "Caja", 2, "Caja", 1400, 2, 0, "")
'    Set grdGrilla.DataSource = RsDetalleAut



End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmCentroCostoDetalle = Nothing
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
