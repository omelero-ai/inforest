VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCajaCodigoControlDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3885
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   11235
   Icon            =   "frmCajaCodigoControlDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   11235
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   11175
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3135
      Width           =   11235
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   6300
         TabIndex        =   23
         Top             =   60
         Width           =   6360
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   4860
            Picture         =   "frmCajaCodigoControlDetalle.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   5340
            Picture         =   "frmCajaCodigoControlDetalle.frx":0984
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   5820
            Picture         =   "frmCajaCodigoControlDetalle.frx":0EC6
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmCajaCodigoControlDetalle.frx":1408
            Style           =   1  'Graphical
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmCajaCodigoControlDetalle.frx":194A
            Style           =   1  'Graphical
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmCajaCodigoControlDetalle.frx":1E8C
            Style           =   1  'Graphical
            TabIndex        =   1
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
            TabIndex        =   24
            Top             =   180
            Width           =   3330
         End
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
         Left            =   7635
         Picture         =   "frmCajaCodigoControlDetalle.frx":23CE
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
         Left            =   6465
         Picture         =   "frmCajaCodigoControlDetalle.frx":2900
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   8805
         Picture         =   "frmCajaCodigoControlDetalle.frx":2E32
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   9975
         Picture         =   "frmCajaCodigoControlDetalle.frx":2F34
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   60
         Width           =   1170
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
      Height          =   3135
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11130
      Begin VB.TextBox txtGlosa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2805
         LinkTimeout     =   10
         MaxLength       =   160
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   1800
         Width           =   7815
      End
      Begin VB.TextBox txtSFC 
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
         Left            =   2805
         LinkTimeout     =   10
         MaxLength       =   100
         TabIndex        =   14
         Top             =   1400
         Width           =   7815
      End
      Begin VB.CheckBox chkActivo 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo :"
         Height          =   195
         Left            =   9780
         TabIndex        =   11
         Top             =   300
         Width           =   840
      End
      Begin VB.TextBox txtAutorizacion 
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
         Left            =   2805
         MaxLength       =   25
         TabIndex        =   12
         Text            =   " "
         Top             =   645
         Width           =   7815
      End
      Begin VB.TextBox txtDosificacion 
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
         Left            =   2805
         MaxLength       =   100
         TabIndex        =   13
         Top             =   1005
         Width           =   7815
      End
      Begin MSComCtl2.DTPicker dtpFechaIni 
         Height          =   315
         Left            =   2805
         TabIndex        =   15
         Top             =   2685
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61931521
         CurrentDate     =   38587
      End
      Begin MSComCtl2.DTPicker dtpFechaFin 
         Height          =   315
         Left            =   8445
         TabIndex        =   16
         Top             =   2685
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61931521
         CurrentDate     =   38587
      End
      Begin MSDataListLib.DataCombo cboCaja 
         Height          =   315
         Left            =   2805
         TabIndex        =   10
         Top             =   255
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         BackColor       =   16777215
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Glosa Para Documentos :"
         Height          =   195
         Left            =   630
         TabIndex        =   28
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblCaja 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6120
         TabIndex        =   27
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Número SFC :"
         Height          =   195
         Left            =   1455
         TabIndex        =   26
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Número de Dosificación :"
         Height          =   195
         Left            =   660
         TabIndex        =   25
         Top             =   1080
         Width           =   1785
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final :"
         Height          =   195
         Index           =   2
         Left            =   7245
         TabIndex        =   21
         Top             =   2760
         Width           =   915
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial :"
         Height          =   195
         Index           =   1
         Left            =   1455
         TabIndex        =   20
         Top             =   2760
         Width           =   990
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Número de Autorización :"
         Height          =   195
         Left            =   660
         TabIndex        =   19
         Top             =   720
         Width           =   1785
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Caja :"
         Height          =   195
         Left            =   2040
         TabIndex        =   18
         Top             =   390
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmCajaCodigoControlDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsCaja As Recordset
Dim fInicio As Date
Dim fFinal As Date
Dim nPrecio As Double
Dim sSemana As String
Dim sql As String
Dim nCorrelativoR As Integer
Dim i As Integer

Sub Asignar()
    With frmCajaCodigoControl.RsCabecera
        'Cuadro de Texto
        txtAutorizacion = IIf(IsNull(!tAutorizacion), "", !tAutorizacion)
        txtDosificacion = IIf(IsNull(!tDosificacion), "", !tDosificacion)
        txtSFC = IIf(IsNull(!tSfc), "", !tSfc)
        txtGlosa = IIf(IsNull(!tGlosaDosificacion), "", !tGlosaDosificacion)
        'Data Combo
        cboCaja.BoundText = IIf(IsNull(!tCaja), "", Trim(!tCaja))
        Me.lblCaja.Caption = IIf(IsNull(!tCaja), "", "- " & cboCaja.BoundText)
        'CHECKBOX
         chkActivo = IIf(IsNull(!lActivo), 0, IIf(!lActivo = True, 1, 0))
        'datapickers
        dtpFechaIni.value = IIf(IsNull(!fInicio), Date, !fInicio)
        dtpFechaFin.value = IIf(IsNull(!fFin), Date, !fFin)
        
        nCorrelativoR = IIf(IsNull(!nCorrelativo), 0, !nCorrelativo)
        
    End With
End Sub

Private Sub cboCaja_Change()
    Me.lblCaja.Caption = "-" & cboCaja.BoundText
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmCajaCodigoControl.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmCajaCodigoControl.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmCajaCodigoControl.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmCajaCodigoControl.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmCajaCodigoControl.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmCajaCodigoControl.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmCajaCodigoControl.RsCabecera.AbsolutePosition & " de " & frmCajaCodigoControl.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   Dim dosificacion As String
   fInicio = Format(dtpFechaIni.value, "yyyy/mm/dd") & " " & Format("00:00:00", "HH:mm:ss")
   fFinal = Format(Me.dtpFechaFin.value, "yyyy/mm/dd") & " " & Format("23:59:59", "HH:mm:ss")
   dosificacion = Replace(Trim(Me.txtDosificacion.Text), "'", "")
   dosificacion = Replace(Trim(dosificacion), ",", "")
   dosificacion = Replace(Trim(dosificacion), """", "")
   dosificacion = Replace(Trim(dosificacion), "´", "")
   Me.txtDosificacion.Text = dosificacion
   Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               Iniciar
               cboCaja.SetFocus
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               
               'chequea datos
               If cboCaja.Text = "" Then MsgBox "Seleccione Caja", vbExclamation, sMensaje: cboCaja.SetFocus: Exit Sub
               If txtAutorizacion.Text = "" Then MsgBox "Ingrese Número de Autorización", vbExclamation, sMensaje: txtAutorizacion.SetFocus: Exit Sub
               If txtDosificacion.Text = "" Then MsgBox "Ingrese Número de Dosificación", vbExclamation, sMensaje: txtDosificacion.SetFocus: Exit Sub
               If txtSFC.Text = "" Then MsgBox "Ingrese Número SFC", vbExclamation, sMensaje: txtSFC.SetFocus: Exit Sub
               If Me.dtpFechaIni.value > Me.dtpFechaFin.value Then MsgBox "Error en rango de Fechas", vbExclamation, sMensaje: dtpFechaIni.SetFocus: Exit Sub
               
               If Sw Then
                   If validaRangoFecha(fInicio, fFinal, cboCaja.BoundText) = True Then
                     'Obtiene el Numero de Orden
                     nCorrelativoR = Calcular("select max(ncorrelativo) as Codigo from TORIGENCODIGOCONTROL where tcaja='" & cboCaja.BoundText & "'", Cn)
                     
                     If IsNull(nCorrelativoR) Or nCorrelativoR = 0 Then
                         nCorrelativoR = 1
                     Else
                         nCorrelativoR = nCorrelativoR + 1
                     End If
                        sql = "insert into TORIGENCODIGOCONTROL (tcaja,ncorrelativo,finicio,ffin,tautorizacion,tdosificacion,tsfc,lactivo,fregistro,tusuario,tGlosaDosificacion) " & _
                                "values ('" & cboCaja.BoundText & "'," & nCorrelativoR & ", '" & Format(fInicio, "yyyy/MM/dd HH:mm:ss") & "','" & Format(fFinal, "yyyy/MM/dd HH:mm:ss") & "','" & Trim(Me.txtAutorizacion.Text) & "','" & Trim(Me.txtDosificacion.Text) & "','" & Trim(Me.txtSFC.Text) & "',1,getdate(),'" & sUsuario & "','" & Trim(txtGlosa.Text) & "')"
                     Cn.Execute sql
                   Else
                    MsgBox "Conflicto en rango de fechas con Registros ya creados", vbCritical, sMensaje
                    Exit Sub
                   End If
                  Sw = False
                  frmCajaCodigoControl.RsCabecera.Sort = "finicio ASC"
                  frmCajaCodigoControl.RsCabecera.Requery
                  frmCajaCodigoControl.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmCajaCodigoControl.RsCabecera.RecordCount = 0, 0, frmCajaCodigoControl.RsCabecera.AbsolutePosition) & " de " & frmCajaCodigoControl.RsCabecera.RecordCount
                  
               Else
                  If validaRangoFechaUpdate(fInicio, fFinal, cboCaja.BoundText, nCorrelativoR) = True Then
                        'Cambiar el SQL
                        sql = " update TORIGENCODIGOCONTROL set tsfc='" & Trim(Me.txtSFC.Text) & "',tautorizacion = '" & Trim(Me.txtAutorizacion.Text) & "' ,tglosadosificacion='" & Trim(txtGlosa.Text) & "', " & _
                                   " tdosificacion='" & Trim(Me.txtDosificacion.Text) & "', finicio='" & Format(fInicio, "yyyy/MM/dd HH:mm:ss") & "', ffin='" & Format(fFinal, "yyyy/MM/dd HH:mm:ss") & "'  " & _
                                    " , Lactivo='" & chkActivo.value & "', fregistro=getdate(),tusuario='" & sUsuario & "' where tcaja='" & cboCaja.BoundText & "' and ncorrelativo=" & nCorrelativoR & " "
                        Cn.Execute sql
                   Else
                        MsgBox "Conflicto en rango de fechas con Registros ya creados", vbCritical, sMensaje
                        Exit Sub
                   End If
                              
                  nPos = frmCajaCodigoControl.RsCabecera.Bookmark
                  frmCajaCodigoControl.RsCabecera.Requery
                  If frmCajaCodigoControl.RsCabecera.RecordCount = 0 Then
                     frmCajaCodigoControl.RsCabecera.Filter = adFilterNone
                  End If
                  frmCajaCodigoControl.RsCabecera.Bookmark = nPos
                  Screen.MousePointer = vbDefault
                  MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
                         
          Case Is = 2 ' Eliminar
               If frmCajaCodigoControl.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               
               If MsgBox("Seguro de Eliminar el Registro ?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
               
               'Cambia el Delete
               Cn.Execute "update TORIGENCODIGOCONTROL set lactivo=0 where tcaja= '" & cboCaja.BoundText & "' and ncorrelativo=" & nCorrelativoR & ""
               frmCajaCodigoControl.RsCabecera.Requery
               If frmCajaCodigoControl.RsCabecera.RecordCount <> 0 Then
                  frmCajaCodigoControl.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmCajaCodigoControl.RsCabecera.RecordCount = 0, 0, frmCajaCodigoControl.RsCabecera.AbsolutePosition) & " de " & frmCajaCodigoControl.RsCabecera.RecordCount
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
    
    'Llena todos los Combos
    LlenaCombos
    
    'Ingrese el SubTitulo
    Me.Caption = " Configuración Número Autorización y Dosificación Detalle "
    
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       Iniciar
   Else
       ActivarBotones (True)
       Asignar
    End If
    cmdTexto.Caption = "Registro " & frmCajaCodigoControl.RsCabecera.AbsolutePosition & " de " & frmCajaCodigoControl.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set RsCaja = Nothing
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


Public Sub Iniciar()
    dtpFechaIni.value = Date
    dtpFechaFin.value = Date
    chkActivo.value = 1
End Sub

Sub LlenaCombos()
    With cboCaja
         Isql = "Select * from tcaja order by tcaja"
         Set RsCaja = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsCaja
             .DataField = "tdescripcion"
             .ListField = "tdescripcion"
             .BoundColumn = "tcaja"
    End With
End Sub
 
Private Sub txtAutorizacion_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtAutorizacion
End Sub

  
Public Function validaRangoFecha(fIni As Date, fFin As Date, ccaja As String) As Boolean
    Dim registroC As Integer
    validaRangoFecha = True
    registroC = Calcular("select count(*) as codigo from torigencodigocontrol where (tcaja='" & ccaja & "') and (((finicio <='" & Format(fIni, "yyyy/MM/dd HH:mm:ss") & "')  and (ffin>='" & Format(fIni, "yyyy/MM/dd HH:mm:ss") & "')) or ((finicio <='" & Format(fFin, "yyyy/MM/dd HH:mm:ss") & "')  and (ffin>='" & Format(fFin, "yyyy/MM/dd HH:mm:ss") & "') ))", Cn)
    If registroC > 0 Then: validaRangoFecha = False: Exit Function
'    registroC = Calcular("select count(*) as codigo from torigencodigocontrol where (tcaja='" & ccaja & "' and finicio >= '" & Format(fIni, "yyyy/MM/dd HH:mm:ss") & "') or ( tcaja='" & ccaja & "' and ffin >='" & Format(fFin, "yyyy/MM/dd HH:mm:ss") & "')", Cn)
'    If registroC > 0 Then: validaRangoFecha = False: Exit Function
End Function

Public Function validaRangoFechaUpdate(fIni As Date, fFin As Date, ccaja As String, cCorre As Integer) As Boolean
    Dim registroC As Integer
    validaRangoFechaUpdate = True
    registroC = Calcular("select count(*) as codigo from torigencodigocontrol where (tcaja='" & ccaja & "' and ncorrelativo<>" & cCorre & ") and (((finicio <='" & Format(fIni, "yyyy/MM/dd HH:mm:ss") & "')  and (ffin>='" & Format(fIni, "yyyy/MM/dd HH:mm:ss") & "')) or ((finicio <='" & Format(fFin, "yyyy/MM/dd HH:mm:ss") & "')  and (ffin>='" & Format(fFin, "yyyy/MM/dd HH:mm:ss") & "') ))", Cn)
    If registroC > 0 Then: validaRangoFechaUpdate = False: Exit Function
'    registroC = Calcular("select count(*) as codigo from torigencodigocontrol where (tcaja='" & ccaja & "' and ncorrelativo<>" & cCorre & " and finicio >= '" & Format(fIni, "yyyy/MM/dd HH:mm:ss") & "') or ( tcaja='" & ccaja & "' and ncorrelativo<>" & cCorre & " and ffin >='" & Format(fFin, "yyyy/MM/dd HH:mm:ss") & "')", Cn)
'    If registroC > 0 Then: validaRangoFechaUpdate = False: Exit Function
End Function

Private Sub txtGlosa_LostFocus()
 Call ValidaStr(txtGlosa)
End Sub

Private Sub txtSFC_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, txtAutorizacion
End Sub
