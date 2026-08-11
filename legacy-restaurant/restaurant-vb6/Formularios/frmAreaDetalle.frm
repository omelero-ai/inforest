VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAreaDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3870
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9450
   Icon            =   "frmAreaDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   9450
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9390
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3120
      Width           =   9450
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
         Picture         =   "frmAreaDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Picture         =   "frmAreaDetalle.frx":0534
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Picture         =   "frmAreaDetalle.frx":0636
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmAreaDetalle.frx":0B68
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4590
         TabIndex        =   15
         Top             =   60
         Width           =   4650
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmAreaDetalle.frx":109A
            Style           =   1  'Graphical
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmAreaDetalle.frx":15DC
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmAreaDetalle.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4110
            Picture         =   "frmAreaDetalle.frx":2060
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3630
            Picture         =   "frmAreaDetalle.frx":25A2
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3150
            Picture         =   "frmAreaDetalle.frx":2AE4
            Style           =   1  'Graphical
            TabIndex        =   16
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
            TabIndex        =   23
            Top             =   150
            Width           =   1665
         End
      End
   End
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
      Height          =   3075
      Left            =   2100
      TabIndex        =   22
      Top             =   30
      Width           =   7305
      Begin VB.CheckBox chkKdsx 
         Alignment       =   1  'Right Justify
         Caption         =   "KDS 2:"
         Height          =   255
         Left            =   2400
         TabIndex        =   28
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox chkCheffControl 
         Alignment       =   1  'Right Justify
         Caption         =   "Imprime en Pantalla :"
         Height          =   255
         Left            =   345
         TabIndex        =   9
         Top             =   2376
         Width           =   1785
      End
      Begin VB.CheckBox chkKds 
         Alignment       =   1  'Right Justify
         Caption         =   "KDS 1:"
         Height          =   255
         Left            =   1185
         TabIndex        =   8
         Top             =   2055
         Width           =   960
      End
      Begin VB.CheckBox chkDefecto 
         Caption         =   "Por defecto"
         Height          =   255
         Left            =   4800
         TabIndex        =   5
         Top             =   1368
         Width           =   1215
      End
      Begin VB.CheckBox chkSalon 
         Alignment       =   1  'Right Justify
         Caption         =   "Punto de Impresión :"
         Height          =   225
         Left            =   390
         TabIndex        =   3
         Top             =   1383
         Width           =   1755
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
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   681
         Width           =   5250
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
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   2
         Text            =   " "
         Top             =   1032
         Width           =   2595
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   330
         Width           =   1170
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   1290
         TabIndex        =   7
         Top             =   2700
         Width           =   855
      End
      Begin MSDataListLib.DataCombo cboSalon 
         Height          =   315
         Left            =   2190
         TabIndex        =   4
         Top             =   1338
         Width           =   2355
         _ExtentX        =   4154
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
      Begin MSDataListLib.DataCombo cboCosto 
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   1674
         Width           =   2595
         _ExtentX        =   4577
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
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   1290
         TabIndex        =   27
         Top             =   375
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Resumida :"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   26
         Top             =   1077
         Width           =   1680
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Detallada :"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   25
         Top             =   726
         Width           =   1650
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo :"
         Height          =   195
         Index           =   3
         Left            =   645
         TabIndex        =   24
         Top             =   1734
         Width           =   1230
      End
   End
   Begin VB.Image Image 
      Height          =   3060
      Left            =   15
      Picture         =   "frmAreaDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   15
      Width           =   2040
   End
End
Attribute VB_Name = "frmAreaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSalon As Recordset
Dim RsCosto As Recordset

Sub LlenaCombo()
    If lAlmacen And CnAlmacen.State Then
        With cboCosto
             Isql = "Select * from vCArea where lActivo = 1"
             Set RsCosto = Lib.OpenRecordset(Isql, CnAlmacen)
             Set .RowSource = RsCosto
                 .DataField = "tDetallado"
                 .ListField = "tDetallado"
                 .BoundColumn = "tCodigo"
        End With
    End If
    
    With cboSalon
         Isql = "Select * from vSalon where lActivo = 1 order by Descripcion"
         Set RsSalon = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsSalon
             .DataField = "Descripcion"
             .ListField = "Descripcion"
             .BoundColumn = "Codigo"
    End With
    
End Sub
Sub Asignar()
    With frmArea.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
        cboSalon.BoundText = IIf(IsNull(!tIcono), "", Trim(!tIcono))
        If cboSalon.Text = "" Then
           cboSalon.Enabled = False
           chkSalon.Value = 0
           chkDefecto.Enabled = False
        Else
           cboSalon.Enabled = True
           chkSalon.Value = 1
           chkDefecto.Enabled = True
        End If
        cboCosto.BoundText = IIf(IsNull(!tvalor), "", Trim(!tvalor))
        
        'Check Box
        chkActivo = IIf(!lActivo = True, 1, 0)
        chkDefecto = IIf(IsNull(!nValor), 0, !nValor)
        
        'KDS
        If (lKDS) Then
            chkKDS = IIf(!kds = 1, 1, 0)
            chkKdsx = IIf(!KDSX = True, 1, 0)
        Else
            chkKDS = 0
            chkKDS.Enabled = False
            chkKdsx = 0
            chkKdsx.Enabled = False
        End If
        
        
        If !lCheffControl = 0 Then
            chkCheffControl.Value = 0
        ElseIf !lCheffControl = 1 Then
            chkCheffControl.Value = 1
        Else
            chkCheffControl.Value = 0
        End If
        
        
    End With
End Sub

Private Sub chkKds_Click()
'    If chkKds.Value Then
'        chkKdsx.Value = False
'    End If
End Sub

Private Sub chkKdsx_Click()
'    If chkKdsx.Value Then
'        chkKds.Value = False
'    End If

End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmArea.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmArea.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmArea.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmArea.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmArea.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmArea.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmArea.RsCabecera.RecordCount = 0, 0, frmArea.RsCabecera.AbsolutePosition) & " de " & frmArea.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               chkActivo.Value = 1
               'Cambia el Nombre del Primer Text
               txtDetallado.SetFocus
                    
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción Detallada", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese la Descripción Resumida", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
                    
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='Area' ", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "001"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 3)
                  End If
                  Sw = False
                                  
                  'Cambiar el SQL 'KDS
                  Isql = "insert into TTABLA( " & _
                         "tTabla, tCodigo, tDetallado, tResumido, tValor, tIcono, nBoton, nValor, lActivo,ntamano,lAutorizaEnviosMozo) " & _
                         "values ('AREA', " & _
                                " '" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                " '" & cboCosto.BoundText & "', " & _
                                " '" & cboSalon.BoundText & "', " & _
                                       chkKDS.Value & ", " & _
                                       chkDefecto.Value & ", " & _
                                       chkActivo.Value & "," & IIf(chkCheffControl.Value = 1, 1, 0) & "," & IIf(chkKdsx.Value = 1, 1, 0) & ") "
                  Cn.Execute Isql
                  ActivarBotones (True)
                  frmArea.RsCabecera.Sort = "Codigo ASC"
                  frmArea.RsCabecera.Requery
                  frmArea.RsCabecera.MoveLast
                  cmdTexto.Caption = "Registro " & IIf(frmArea.RsCabecera.RecordCount = 0, 0, frmArea.RsCabecera.AbsolutePosition) & " de " & frmArea.RsCabecera.RecordCount
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  
               Else
                  'Cambiar el SQL 'KDS
                  Isql = "update TTABLA set " & _
                         "tDetallado ='" & txtDetallado.Text & "', " & _
                         "tResumido ='" & txtResumido.Text & "', " & _
                         "tValor ='" & cboCosto.BoundText & "', " & _
                         "tIcono ='" & cboSalon.BoundText & "', " & _
                         "nBoton =" & chkKDS.Value & ", " & _
                         "lAutorizaEnviosMozo =" & chkKdsx.Value & ", " & _
                         "nValor =" & chkDefecto.Value & "," & _
                         "lActivo =" & chkActivo.Value & ", lReplica=1, ntamano=" & IIf(chkCheffControl.Value = 1, 1, 0) & " " & _
                         " where tTAbla = 'AREA' and tCodigo = '" & txtCodigo & "'"
                   Cn.Execute Isql
                   nPos = frmArea.RsCabecera.Bookmark
                   frmArea.RsCabecera.Requery
                   If frmArea.RsCabecera.RecordCount = 0 Then
                      frmArea.RsCabecera.Filter = adFilterNone
                   End If
                   frmArea.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
               Cn.Execute "update TTABLA set  lNuevoModificado=1  where tTAbla = 'AREA' and tCodigo = '" & txtCodigo & "'"
          Case Is = 2 ' Eliminar
               If frmArea.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Area " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "delete from TTABLA where tTabla = 'AREA' and tCodigo = '" & txtCodigo & "'"
               Cn.Execute "delete from TPRODUCTOAREA where tArea='" & txtCodigo & "'"
               frmArea.RsCabecera.Requery
               If frmArea.RsCabecera.RecordCount <> 0 Then
                  frmArea.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmArea.RsCabecera.RecordCount = 0, 0, frmArea.RsCabecera.AbsolutePosition) & " de " & frmArea.RsCabecera.RecordCount
               Else
                  ActivarBotones False
                  Blanquear Me
                  Sw = True
               End If
          
          Case Is = 3 ' Salir
               Unload Me
   End Select

End Sub

Private Sub chkSalon_Click()
   If chkSalon.Value = 1 Then
      cboSalon.Enabled = True
      chkDefecto.Enabled = True
   Else
      cboSalon.Enabled = False
      cboSalon.Text = ""
      chkDefecto.Enabled = False
      chkDefecto.Value = 0
   End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Centrar Me
        
    'Ingrese el SubTitulo
    Me.Caption = " Areas de Producción"
    fraDetalle.Caption = Me.Caption
    LlenaCombo
    'Ingresar la Tabla
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.Value = 1
       chkSalon.Value = 0
       cboSalon.Enabled = False
    Else
       ActivarBotones (True)
       Asignar
    End If
    cmdTexto.Caption = "Registro " & IIf(frmArea.RsCabecera.RecordCount = 0, 0, frmArea.RsCabecera.AbsolutePosition) & " de " & frmArea.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAreaDetalle = Nothing
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
   Call ValidaStr(txtDetallado)
End Sub

Private Sub txtResumido_LostFocus()
   Call ValidaStr(txtResumido)
End Sub
