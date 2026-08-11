VERSION 5.00
Begin VB.Form frmTipoIdentidadDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4860
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9480
   Icon            =   "frmTipoIdentidadDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   9480
   Begin VB.TextBox txtCodigoIdentidad 
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
      Left            =   4230
      MaxLength       =   15
      TabIndex        =   22
      Text            =   " "
      Top             =   1440
      Width           =   1215
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
      Height          =   4080
      Left            =   2280
      TabIndex        =   17
      Top             =   0
      Width           =   7155
      Begin VB.CheckBox ChkSolicitaDatosSeparados 
         Alignment       =   1  'Right Justify
         Caption         =   "Solicitar datos separados (Nombre, Segundo Nombre, Apellido, Segundo Apellido) "
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   3360
         TabIndex        =   38
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Frame Frame1 
         Height          =   1335
         Left            =   120
         TabIndex        =   34
         Top             =   2640
         Width           =   3495
         Begin VB.TextBox txtMonto 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2280
            TabIndex        =   36
            Text            =   "0"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chkActivaMonto 
            Caption         =   "Excluir de todo proceso de facturacion con monto mayor:"
            Height          =   675
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "* Es necesario activar en Tipo de Documento 'solicita cliente deacuerdo al monto'"
            ForeColor       =   &H000040C0&
            Height          =   375
            Left            =   120
            TabIndex        =   37
            Top             =   840
            Width           =   3255
         End
      End
      Begin VB.Frame frmValidacionTipo 
         Caption         =   "TIpo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   5280
         TabIndex        =   30
         Top             =   2640
         Visible         =   0   'False
         Width           =   1815
         Begin VB.OptionButton OptAlfanumerico 
            Caption         =   "Alfanumerico"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton OptNumerico 
            Caption         =   "Numerico"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Frame frmValidacionLongitud 
         Caption         =   "Longitud"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   3600
         TabIndex        =   26
         Top             =   2640
         Visible         =   0   'False
         Width           =   1695
         Begin VB.OptionButton optMenorIgual 
            Caption         =   "<= Longitud"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Width           =   1335
         End
         Begin VB.OptionButton OptIgual 
            Caption         =   "  = Longitud"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1080
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton optMayorIgual 
            Caption         =   ">= Longitud"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   840
            Width           =   1335
         End
         Begin VB.TextBox txtLongitud 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            MaxLength       =   4
            TabIndex        =   27
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtRefTipoPersona 
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
         Left            =   5235
         MaxLength       =   15
         TabIndex        =   25
         Text            =   " "
         Top             =   1440
         Width           =   1770
      End
      Begin VB.CheckBox chkValidacion 
         Alignment       =   1  'Right Justify
         Caption         =   "Validación :"
         Height          =   195
         Left            =   1020
         TabIndex        =   3
         Top             =   1830
         Width           =   1125
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
         Left            =   1950
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   712
         Width           =   5070
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
         Left            =   1950
         MaxLength       =   15
         TabIndex        =   2
         Text            =   " "
         Top             =   1094
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
         Left            =   1950
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
         Left            =   1305
         TabIndex        =   4
         Top             =   2130
         Width           =   840
      End
      Begin VB.Label lblRefTipoPersona 
         AutoSize        =   -1  'True
         Caption         =   "Referencia Tipo Persona :"
         Height          =   195
         Left            =   3315
         TabIndex        =   24
         Top             =   1480
         Width           =   1860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código Identidad :"
         Height          =   195
         Left            =   580
         TabIndex        =   23
         Top             =   1480
         Width           =   1290
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Detallada :"
         Height          =   195
         Left            =   210
         TabIndex        =   20
         Top             =   757
         Width           =   1650
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Descripción Resumida :"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   1139
         Width           =   1680
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   1275
         TabIndex        =   18
         Top             =   375
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9420
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4110
      Width           =   9480
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
         Picture         =   "frmTipoIdentidadDetalle.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmTipoIdentidadDetalle.frx":0534
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
         Left            =   4710
         Picture         =   "frmTipoIdentidadDetalle.frx":0636
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "frmTipoIdentidadDetalle.frx":0B68
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4590
         TabIndex        =   10
         Top             =   60
         Width           =   4650
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmTipoIdentidadDetalle.frx":109A
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmTipoIdentidadDetalle.frx":15DC
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmTipoIdentidadDetalle.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4110
            Picture         =   "frmTipoIdentidadDetalle.frx":2060
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3630
            Picture         =   "frmTipoIdentidadDetalle.frx":25A2
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3150
            Picture         =   "frmTipoIdentidadDetalle.frx":2AE4
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
            Left            =   1440
            TabIndex        =   21
            Top             =   150
            Width           =   1665
         End
      End
   End
   Begin VB.Image Image 
      Height          =   4050
      Left            =   45
      Picture         =   "frmTipoIdentidadDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   15
      Width           =   2205
   End
End
Attribute VB_Name = "frmTipoIdentidadDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Asignar()
    With frmTipoIdentidad.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!tResumido), "", !tResumido)
            
        'Check Box
        chkValidacion = IIf(IsNull(!nValor), 0, !nValor)
        chkActivo = IIf(!lActivo = True, 1, 0)
        
        txtCodigoIdentidad = IIf(IsNull(!tValor), "", !tValor)
        
        txtRefTipoPersona = IIf(IsNull(!tValor2), "", !tValor2)
        
        txtLongitud.Text = IIf(IsNull(!nLongitud), "", !nLongitud)
        
        Me.optMayorIgual.Value = IIf(!lMayorIgual, 1, 0)
        Me.optMenorIgual.Value = IIf(!lMenorIgual, 1, 0)
        Me.OptIgual.Value = IIf(!lIgual, 1, 0)
        
        
        Me.OptAlfanumerico.Value = IIf(!lTipoDato, 0, 1)
        Me.OptNumerico.Value = IIf(!lTipoDato, 1, 0)
        
        If (chkValidacion.Value = 0) Then
            Me.frmValidacionLongitud.Visible = False
            Me.frmValidacionTipo.Visible = False
        End If
        
        Me.chkActivaMonto.Value = IIf(!lactivamonto, 1, 0)
        Me.txtMonto.Text = IIf(IsNull(!nMonto), "0", !nMonto)
        Me.ChkSolicitaDatosSeparados.Value = IIf(!lActivaDatosSeparados = True, 1, 0)
        
    End With
End Sub

Private Sub chkValidacion_Click()
    If chkValidacion.Value = 1 Then
        Me.frmValidacionLongitud.Visible = True
        Me.frmValidacionTipo.Visible = True
    Else
        Me.frmValidacionLongitud.Visible = False
        Me.frmValidacionTipo.Visible = False
    End If
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmTipoIdentidad.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmTipoIdentidad.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmTipoIdentidad.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmTipoIdentidad.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmTipoIdentidad.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmTipoIdentidad.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmTipoIdentidad.RsCabecera.AbsolutePosition & " de " & frmTipoIdentidad.RsCabecera.RecordCount
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
                  'nCorrela = Calcular("select max(tCodigo) as Codigo from TTABLA where tTabla ='TIPOIDENTIDAD' ", Cn)
                  nCorrela = Calcular("select max(tCodigo) as Codigo from TTIPOIDENTIDAD", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "01"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 2)
                  End If
                  Sw = False
                   
                  'Cambiar el SQL
'                  Isql = "insert into TTABLA( " & _
'                         "tTabla, tCodigo, tDetallado, tResumido, nValor, tValor, tValor2, lActivo) " & _
'                         "values ('TIPOIDENTIDAD', " & _
'                                " '" & txtCodigo.Text & "', " & _
'                                " '" & txtDetallado.Text & "', " & _
'                                " '" & txtResumido.Text & "', " & _
'                                       chkValidacion.value & ", " & _
'                                " '" & txtCodigoIdentidad.Text & "', " & _
'                                " '" & txtRefTipoPersona.Text & "', " & _
'                                       chkActivo.value & ") "
                                       
                  Isql = "insert into TTIPOIDENTIDAD( " & _
                         " tCodigo, Descripcion, tResumido, nValor, tValor, tValor2, lActivo, nLongitud, lMayorIgual,lTipoDato, lMenorIgual,lIgual, lActivaMonto, nMonto, lactivaDatosSeparados) " & _
                         "values ('" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                       chkValidacion.Value & ", " & _
                                " '" & txtCodigoIdentidad.Text & "', " & _
                                " '" & txtRefTipoPersona.Text & "', " & _
                                       chkActivo.Value & "," & val(txtLongitud.Text) & ", " & _
                                       IIf(Me.optMayorIgual.Value, 1, 0) & "," & IIf(Me.OptNumerico.Value, 1, 0) & "," & _
                                       IIf(Me.optMenorIgual.Value, 1, 0) & "," & IIf(Me.OptIgual.Value, 1, 0) & ", " & _
                                        IIf(Me.chkActivaMonto.Value, 1, 0) & ", " & CDbl(Me.txtMonto.Text) & ", " & IIf(Me.chkActivaMonto.Value, 1, 0) & ") "
                    
           
                      Cn.Execute Isql
                      frmTipoIdentidad.RsCabecera.Sort = "Codigo ASC"
                      frmTipoIdentidad.RsCabecera.Requery
                      frmTipoIdentidad.RsCabecera.MoveLast
                      MsgBox "Registro Guardado", vbInformation, sMensaje
                      ActivarBotones (True)
                      cmdTexto.Caption = "Registro " & IIf(frmTipoIdentidad.RsCabecera.RecordCount = 0, 0, frmTipoIdentidad.RsCabecera.AbsolutePosition) & " de " & frmTipoIdentidad.RsCabecera.RecordCount
              Else
                 'Cambiar el SQL
'                 Isql = "update TTABLA set " & _
'                        "tDetallado ='" & txtDetallado.Text & "', " & _
'                        "tResumido ='" & txtResumido.Text & "', " & _
'                        "nValor =" & chkValidacion.value & ", " & _
'                        "tValor ='" & txtCodigoIdentidad.Text & "', " & _
'                        "tValor2 ='" & txtRefTipoPersona.Text & "', " & _
'                        "lActivo =" & chkActivo.value & _
'                        " where tTAbla = 'TIPOIDENTIDAD' and tCodigo = '" & txtCodigo & "'"
                 
                 Isql = "update TTIPOIDENTIDAD set " & _
                        "Descripcion ='" & txtDetallado.Text & "', " & _
                        "tResumido ='" & txtResumido.Text & "', " & _
                        "nValor =" & chkValidacion.Value & ", " & _
                        "tValor ='" & txtCodigoIdentidad.Text & "', " & _
                        "tValor2 ='" & txtRefTipoPersona.Text & "', " & _
                        "lActivo =" & chkActivo.Value & ", " & _
                        "nLongitud=" & val(txtLongitud.Text) & ", " & _
                        "lMayorIgual= " & IIf(Me.optMayorIgual.Value, 1, 0) & ", " & _
                        "lMenorIgual= " & IIf(Me.optMenorIgual.Value, 1, 0) & ", " & _
                        "lIgual= " & IIf(Me.OptIgual.Value, 1, 0) & ", " & _
                        "lTipoDato= " & IIf(Me.OptNumerico.Value, 1, 0) & ",  " & _
                        "lActivaMonto=" & IIf(Me.chkActivaMonto.Value, 1, 0) & ", " & _
                        "nMonto= " & CDbl(Me.txtMonto.Text) & ", " & _
                        "lActivaDatosSeparados= " & Me.ChkSolicitaDatosSeparados.Value & " " & _
                        " where tCodigo = '" & txtCodigo & "'"
                        
                  Cn.Execute Isql
                  nPos = frmTipoIdentidad.RsCabecera.Bookmark
                  frmTipoIdentidad.RsCabecera.Requery
                  If frmTipoIdentidad.RsCabecera.RecordCount = 0 Then
                     frmTipoIdentidad.RsCabecera.Filter = adFilterNone
                  End If
                  frmTipoIdentidad.RsCabecera.Bookmark = nPos
                  Screen.MousePointer = vbDefault
                  MsgBox "Registro Modificado", vbInformation, sMensaje
              End If
              
            Cn.Execute "update TTIPOIDENTIDAD set lNuevoModificado=1 where tCodigo = '" & txtCodigo & "'"
          
          Case Is = 2 ' Eliminar
               If frmTipoIdentidad.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Tipo de Identidad, Esto puede traer errores en el sistema!!!, Codigo: " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "delete from TTIPOIDENTIDAD where tCodigo = '" & txtCodigo & "'"
               frmTipoIdentidad.RsCabecera.Requery
               If frmTipoIdentidad.RsCabecera.RecordCount <> 0 Then
                  frmTipoIdentidad.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmTipoIdentidad.RsCabecera.RecordCount = 0, 0, frmTipoIdentidad.RsCabecera.AbsolutePosition) & " de " & frmTipoIdentidad.RsCabecera.RecordCount
                  
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
    Screen.MousePointer = vbHourglass
    Centrar Me
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Tipo de Identidad "
    fraDetalle.Caption = Me.Caption
       
    'Ingresar la Tabla
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.Value = 1
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    
    If lSAP Then
        lblRefTipoPersona.Visible = True
        txtRefTipoPersona.Visible = True
    Else
        lblRefTipoPersona.Visible = False
        txtRefTipoPersona.Visible = False
    End If
    
    cmdTexto.Caption = "Registro " & frmTipoIdentidad.RsCabecera.AbsolutePosition & " de " & frmTipoIdentidad.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTipoIdentidadDetalle = Nothing
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



Private Sub txtMonto_Change()
On Error GoTo fin
If Trim(txtMonto.Text) = "" Then txtMonto = "0": Exit Sub
If Not IsNumeric(txtMonto.Text) Then
    MsgBox "El Valor ingresado debe ser un numero!!!"
    txtMonto.Text = "0"
End If

Exit Sub
fin:
End Sub

Private Sub txtLongitud_Change()
On Error GoTo fin
If Trim(txtLongitud.Text) = "" Then txtLongitud = "0": Exit Sub
If Not IsNumeric(txtLongitud.Text) Then
    MsgBox "El Valor ingresado debe ser un numero!!!"
    txtLongitud.Text = "0"
End If

Exit Sub
fin:
End Sub
