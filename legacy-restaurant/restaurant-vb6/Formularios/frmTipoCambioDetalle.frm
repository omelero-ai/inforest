VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTipoCambioDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo Cambio"
   ClientHeight    =   3480
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   9735
   Icon            =   "frmTipoCambioDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   9735
   Begin VB.Frame fraDetalle 
      Caption         =   "Tipo Cambio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      Left            =   2370
      TabIndex        =   21
      Top             =   30
      Width           =   7320
      Begin VB.Frame frmCambio 
         BorderStyle     =   0  'None
         Height          =   1020
         Left            =   20
         TabIndex        =   29
         Top             =   640
         Visible         =   0   'False
         Width           =   5415
         Begin VB.TextBox txtVenta2 
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
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   31
            Text            =   " "
            Top             =   380
            Width           =   1560
         End
         Begin VB.TextBox txtVenta1 
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
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   30
            Text            =   " "
            Top             =   60
            Width           =   1560
         End
         Begin VB.TextBox txtVenta3 
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
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   33
            Text            =   " "
            Top             =   680
            Width           =   1560
         End
         Begin VB.Label lblcambio2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio Venta 2 :"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   405
            Width           =   1815
         End
         Begin VB.Label lblCambio1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio Venta 1 :"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label lblCambio3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cambio Venta 3 :"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   705
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Oficial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   810
      End
      Begin MSComCtl2.DTPicker DTFecha 
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61538305
         CurrentDate     =   41019
      End
      Begin VB.TextBox txtTCSunat 
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   11
         Text            =   " "
         Top             =   1350
         Width           =   1560
      End
      Begin VB.TextBox txtCompra 
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   7
         Text            =   " "
         Top             =   690
         Width           =   1560
      End
      Begin VB.TextBox txtVenta 
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
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   9
         Text            =   " "
         Top             =   1020
         Width           =   1560
      End
      Begin VB.TextBox txtRegistro 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   2070
         Width           =   2160
      End
      Begin VB.TextBox txtUsuario 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1710
         Width           =   2160
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   810
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cambio Oficial :"
         Height          =   195
         Left            =   360
         TabIndex        =   28
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cambio Compra :"
         Height          =   195
         Left            =   330
         TabIndex        =   26
         Top             =   720
         Width           =   1560
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cambio Venta :"
         Height          =   195
         Left            =   450
         TabIndex        =   25
         Top             =   1065
         Width           =   1440
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   1350
         TabIndex        =   24
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Usuario :"
         Height          =   195
         Left            =   1260
         TabIndex        =   23
         Top             =   1770
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Registro :"
         Height          =   195
         Left            =   495
         TabIndex        =   22
         Top             =   2115
         Width           =   1395
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9675
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2730
      Width           =   9735
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
         Left            =   8430
         Picture         =   "frmTipoCambioDetalle.frx":0442
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
         Left            =   7260
         Picture         =   "frmTipoCambioDetalle.frx":0534
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
         Left            =   4920
         Picture         =   "frmTipoCambioDetalle.frx":0636
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Left            =   6090
         Picture         =   "frmTipoCambioDetalle.frx":0B68
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
         ScaleWidth      =   4800
         TabIndex        =   20
         Top             =   60
         Width           =   4860
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmTipoCambioDetalle.frx":109A
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmTipoCambioDetalle.frx":15DC
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
            Picture         =   "frmTipoCambioDetalle.frx":1B1E
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4320
            Picture         =   "frmTipoCambioDetalle.frx":2060
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3840
            Picture         =   "frmTipoCambioDetalle.frx":25A2
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3360
            Picture         =   "frmTipoCambioDetalle.frx":2AE4
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
            Left            =   1410
            TabIndex        =   27
            Top             =   150
            Width           =   1815
         End
      End
   End
   Begin VB.Image Image 
      Height          =   2580
      Left            =   75
      Picture         =   "frmTipoCambioDetalle.frx":3026
      Stretch         =   -1  'True
      Top             =   75
      Width           =   2235
   End
End
Attribute VB_Name = "frmTipoCambioDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Asignar()
    With frmTipoCambio.RsCabecera
        'Cuadro de Texto ---txtCodigo.Text
        DTFecha.value = IIf(IsNull(!fFecha), "", !fFecha)
        txtCompra.Text = IIf(IsNull(!nCompra), "0.000", Format(!nCompra, "###,###,##0.000"))
        txtVenta.Text = IIf(IsNull(!nVenta), "0.000", Format(!nVenta, "###,###,##0.000"))
        txtUsuario.Text = IIf(IsNull(!tUsuario), sUsuario, !tUsuario)
        txtRegistro.Text = IIf(IsNull(!fRegistro), Date & " " & Time, !fRegistro)
        '--CESAR tipo cambio sunat
        txtTCSunat.Text = IIf(IsNull(!nOficial), "0.000", Format(!nOficial, "###,###,##0.000"))
        
        txtVenta1.Text = IIf(IsNull(!nVenta), "0.000", Format(!nVenta, "###,###,##0.000"))
        txtVenta2.Text = IIf(IsNull(!nVenta2), "0.000", Format(!nVenta2, "###,###,##0.000"))
        txtVenta3.Text = IIf(IsNull(!nVenta3), "0.000", Format(!nVenta3, "###,###,##0.000"))
        
    End With
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Agregar
                
                Sw = True
                'Asignar El Campo de Codificación
                ActivarBotones (False)
                Blanquear Me
                frmTipoCambioDetalle.DTFecha.Enabled = True
                'txtCodigo.Text = Date
                DTFecha.value = Date
                'Cambia el Nombre del Primer Text
                txtCompra.SetFocus
                
           Case Is = 1 'Grabar
                'Chequea Datos
                
                
                If txtVenta.Text = "" Or Val(txtVenta) = 0 Then MsgBox "Ingrese el tipo de cambio de la Venta", vbExclamation, sMensaje: txtVenta.SetFocus: Exit Sub
                '----------CESAR tipo cambio suntat
                If pais <> "003" Then
                    If txtCompra.Text = "" Or Val(txtCompra) = 0 Then MsgBox "Ingrese el Tipo de cambio de la Compra", vbExclamation, sMensaje: txtCompra.SetFocus: Exit Sub
                    If txtTCSunat.Text = "" Or Val(txtTCSunat) = 0 Then MsgBox "Ingrese el tipo de cambio oficial", vbExclamation, sMensaje: txtTCSunat.SetFocus: Exit Sub
                End If
                If Sw Then
                   Sw = False
                    'Busca si existe ya registrado la fecha
                    If frmTipoCambio.RsCabecera.RecordCount > 0 Then
                       frmTipoCambio.RsCabecera.Requery
                       frmTipoCambio.RsCabecera.MoveFirst
                       '---CESAR----txtCodigo.Text
                       frmTipoCambio.RsCabecera.Find "fFecha = #" & Format(DTFecha.value, "dd/mm/yyyy") & "#"
                       If Not frmTipoCambio.RsCabecera.EOF Then
                       'CESAR ----sunat txtCodigo.Text
                          MsgBox "El Tipo de Cambio de Fecha " & DTFecha.value & " ya fue registrado", vbCritical, sMensaje
                          Asignar
                          ActivarBotones (True)
                          Exit Sub
                       End If
                    End If
                    
                                sPasa = Format(DTFecha.value, "yyyy/mm/dd")
                    
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPOCAMBIO", "TIPO DE CAMBIO", "01", sUsuario, sPasa, "", _
                "FFECHA", "Fecha", sPasa, "nCompra", "Compra", Val(txtCompra.Text), "nVenta", "Venta", Val(txtVenta.Text), "tUsuario", "Usuario", sUsuario, "nOficial", "Oficial", Val(Me.txtTCSunat.Text))
              If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                    
                    'Cambiar el SQL ---{fn curdate()}
                    Isql = "insert into TTIPOCAMBIO( " & _
                           "fFecha, nCompra, nVenta, nVenta2, nVenta3, tUsuario, nOficial, fRegistro) " & _
                           "values ('" & Format(DTFecha.value, "MM/dd/yyyy") & "', " & _
                                    Val(txtCompra) & ", " & _
                                    IIf(pais = "003", Val(txtVenta1), Val(txtVenta)) & ", " & _
                                    Val(txtVenta2) & ", " & _
                                    Val(txtVenta3) & ", " & _
                                    "'" & sUsuario & "'," & _
                                    IIf(pais = "003", Val(txtVenta1), Val(txtTCSunat)) & ", " & _
                                    " getdate() )"
                    Cn.Execute Isql
                    '
                    frmTipoCambio.RsCabecera.Requery
                    frmTipoCambio.RsCabecera.MoveLast
                    MsgBox "Registro Agregado", vbInformation, sMensaje
                    ActivarBotones (True)
                    cmdTexto.Caption = "Registro " & frmTipoCambio.RsCabecera.AbsolutePosition & " de " & frmTipoCambio.RsCabecera.RecordCount
                    txtUsuario.Text = sUsuario
                    txtRegistro.Text = Date
                    nTipoCambio = txtVenta
                    Asignar
                Else
                
                sPasa = Format(DTFecha.value, "yyyy/mm/dd")
                
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPOCAMBIO", "TIPO DE CAMBIO", "02", sUsuario, sPasa, "", _
                "FFECHA", "Fecha", sPasa, "nCompra", "Compra", Val(txtCompra.Text), "nVenta", "Venta", Val(txtVenta.Text), "tUsuario", "Usuario", sUsuario, "nOficial", "Oficial", Val(Me.txtTCSunat.Text))
              If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                
                
                   'Cambiar el SQL
                   'CESAR----txtCoidgo.Text
                   Isql = "update TTIPOCAMBIO set " & _
                          "nCompra =" & Val(txtCompra.Text) & ", " & _
                          "nVenta =" & Val(txtVenta.Text) & ", " & _
                          "nVenta2 =" & Val(txtVenta2.Text) & ", " & _
                          "nVenta3 =" & Val(txtVenta3.Text) & ", " & _
                          "tUsuario = '" & sUsuario & "', " & _
                          "nOficial =" & Val(txtTCSunat.Text) & ", " & _
                          "fRegistro = getdate() " & _
                          " where fFecha = '" & Format(DTFecha.value, "yyyy/mm/dd") & "'"
                          
                    Cn.Execute Isql
                                                                        
                   nPos = frmTipoCambio.RsCabecera.Bookmark
                   frmTipoCambio.RsCabecera.Requery
                   If frmTipoCambio.RsCabecera.RecordCount = 0 Then
                      frmTipoCambio.RsCabecera.Filter = adFilterNone
                   End If
                   frmTipoCambio.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
                    
                   nTipoCambio = txtVenta
                   txtUsuario.Text = sUsuario
                   txtRegistro.Text = Format(Date + Time, "dd/mm/yyyy hh:mm AMPM")
                   Asignar
                End If
           
           Case Is = 2 ' Eliminar
                If frmTipoCambio.RsCabecera.RecordCount = 0 Then
                   Exit Sub
                End If
                'Cambia el MsgBox
                If MsgBox("Seguro de Eliminar el Tipo de Cambio del día " & DTFecha & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                   Exit Sub
                End If
                
                
                sPasa = Format(DTFecha.value, "yyyy/mm/dd")
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPOCAMBIO", "TIPO DE CAMBIO", "02", sUsuario, sPasa, "", _
                "FFECHA", "Fecha", sPasa, "nCompra", "Compra", Val(txtCompra.Text), "nVenta", "Venta", Val(txtVenta.Text), "tUsuario", "Usuario", sUsuario, "nOficial", "Oficial", Val(Me.txtTCSunat.Text))
                If lAuditoria = False Then
                      Screen.MousePointer = vbDefault
                      Exit Sub
                  End If
                
                                
                
                'Cambia el Delete
                Cn.Execute "delete from TTIPOCAMBIO where fFecha = '" & Format(DTFecha, "yyyy/mm/dd") & "'"
                frmTipoCambio.RsCabecera.Requery
                If frmTipoCambio.RsCabecera.RecordCount <> 0 Then
                   frmTipoCambio.RsCabecera.MoveLast
                   Asignar
                   cmdTexto.Caption = "Registro " & frmTipoCambio.RsCabecera.AbsolutePosition & " de " & frmTipoCambio.RsCabecera.RecordCount
                Else
                   ActivarBotones False
                   Blanquear Me
                   Sw = True
                End If
           
           Case Is = 3
                Unload Me
                
           Case Is = 4 'Compra
                sTipo = "TC"
                frmNumPad.Show vbModal
                txtCompra.Text = IIf(wEnter = True, Format(sDescrip, "###,###,###,##0.000"), txtCompra.Text)

           Case Is = 5 'Venta
                sTipo = "TC"
                frmNumPad.Show vbModal
                txtVenta.Text = IIf(wEnter = True, Format(sDescrip, "###,###,##0.000"), txtVenta.Text)
                If Val(txtVenta.Text) > 0 Then
                   nTipoCambio = Val(txtVenta.Text)
                End If
          Case Is = 6 '' oficial
                sTipo = "TC"
                frmNumPad.Show vbModal
                txtTCSunat.Text = IIf(wEnter = True, Format(sDescrip, "###,###,###,##0.000"), txtTCSunat.Text)
                If Val(txtTCSunat.Text) > 0 Then
                   nTCO = Val(txtTCSunat.Text)
                End If
                  
    End Select
End Sub

Private Sub Form_Load()
    Centrar Me
    'Ingrese el SubTitulo
    sTitulo = " Mantenimiento de Tipo de Cambio"
    fraDetalle.Caption = sTitulo
    Me.Caption = sTitulo
    
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       DTFecha.value = Date
    Else
       ActivarBotones (True)
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & frmTipoCambio.RsCabecera.AbsolutePosition & " de " & frmTipoCambio.RsCabecera.RecordCount
    
    If pais = "003" Then
        cmdOpcion(4).Visible = False
        cmdOpcion(5).Visible = False
        cmdOpcion(6).Visible = False
        frmCambio.Visible = True
        
        Me.lblCambio1.Caption = "Tipo Cambio Venta " & Calcular("Select isnull(tresumido,'') as codigo from vmoneda where codigo='02'", Cn)
        Me.lblcambio2.Caption = "Tipo Cambio Venta " & Calcular("Select isnull(tresumido,'') as codigo from vmoneda where codigo='03'", Cn)
        Me.lblCambio3.Caption = "Tipo Cambio Venta " & Calcular("Select isnull(tresumido,'') as codigo from vmoneda where codigo='04'", Cn)
        
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmTipoCambioDetalle = Nothing
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmTipoCambio.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmTipoCambio.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmTipoCambio.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmTipoCambio.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmTipoCambio.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmTipoCambio.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & frmTipoCambio.RsCabecera.AbsolutePosition & " de " & frmTipoCambio.RsCabecera.RecordCount
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

Private Sub txtVenta1_Change()
Me.txtVenta.Text = Me.txtVenta1.Text
End Sub

Private Sub txtVenta1_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, Me.txtVenta1.Text
 'Me.txtVenta.Text = Me.txtVenta1.Text
End Sub

Private Sub txtVenta2_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, Me.txtVenta2.Text
End Sub
Private Sub txtVenta3_KeyPress(KeyAscii As Integer)
   TabNext KeyAscii
   Numerico KeyAscii, Me.txtVenta3.Text
End Sub
