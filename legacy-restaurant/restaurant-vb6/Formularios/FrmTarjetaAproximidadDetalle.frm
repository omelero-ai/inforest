VERSION 5.00
Begin VB.Form FrmTarjetaAproximidadDetalle 
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDetalle 
      Height          =   2535
      Left            =   2250
      TabIndex        =   12
      Top             =   0
      Width           =   7220
      Begin VB.CommandButton BtnBuscar 
         Caption         =   "Buscar Cliente"
         Height          =   375
         Left            =   3360
         TabIndex        =   24
         Top             =   1120
         Width           =   1455
      End
      Begin VB.ComboBox CbEstado 
         Height          =   315
         Left            =   1680
         TabIndex        =   21
         Top             =   2080
         Width           =   2775
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   240
         Width           =   5415
      End
      Begin VB.TextBox txtDetallado 
         Height          =   285
         Left            =   1680
         TabIndex        =   15
         Top             =   550
         Width           =   5415
      End
      Begin VB.TextBox txtMonto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   870
         Width           =   1575
      End
      Begin VB.TextBox txtCodCliente 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   1180
         Width           =   1575
      End
      Begin VB.Label LBLID 
         Caption         =   "Label5"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cliente : "
         Height          =   195
         Left            =   1000
         TabIndex        =   25
         Top             =   1640
         Width           =   615
      End
      Begin VB.Label LblCliente 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1680
         TabIndex        =   23
         Top             =   1540
         Width           =   5415
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Estado : "
         Height          =   195
         Left            =   1005
         TabIndex        =   22
         Top             =   2150
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Left            =   1000
         TabIndex        =   20
         Top             =   250
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
         Height          =   195
         Left            =   1000
         TabIndex        =   19
         Top             =   570
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Monto Disponible :"
         Height          =   195
         Left            =   280
         TabIndex        =   18
         Top             =   880
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Codigo del Cliente :"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1200
         Width           =   1365
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   9420
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2580
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
         Left            =   7200
         Picture         =   "FrmTarjetaAproximidadDetalle.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   4830
         Picture         =   "FrmTarjetaAproximidadDetalle.frx":00F2
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
         Left            =   6000
         Picture         =   "FrmTarjetaAproximidadDetalle.frx":0624
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4590
         TabIndex        =   1
         Top             =   60
         Width           =   4650
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "FrmTarjetaAproximidadDetalle.frx":0B56
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "FrmTarjetaAproximidadDetalle.frx":1098
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "FrmTarjetaAproximidadDetalle.frx":15DA
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4110
            Picture         =   "FrmTarjetaAproximidadDetalle.frx":1B1C
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3630
            Picture         =   "FrmTarjetaAproximidadDetalle.frx":205E
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3150
            Picture         =   "FrmTarjetaAproximidadDetalle.frx":25A0
            Style           =   1  'Graphical
            TabIndex        =   2
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
            TabIndex        =   8
            Top             =   150
            Width           =   1665
         End
      End
   End
   Begin VB.Image Image 
      Height          =   2535
      Left            =   0
      Picture         =   "FrmTarjetaAproximidadDetalle.frx":2AE2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2205
   End
End
Attribute VB_Name = "FrmTarjetaAproximidadDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nMonto As Double

Private Sub BtnBuscar_Click()
sTemp = ""
Isql = "select  tCodigoDelivery codigo , tIdentidad , tApellido + ' ' + tNombre descripcion   from TDELIVERY"
frmBusquedaRapida.cmdOpcion(1).Enabled = False
frmBusquedaRapida.cmdOpcion(2).Enabled = False
frmBusquedaRapida.cmdOpcion(3).Enabled = False
frmBusquedaRapida.cmdOpcion(1).Visible = False
frmBusquedaRapida.cmdOpcion(2).Visible = False
frmBusquedaRapida.cmdOpcion(3).Visible = False
frmBusquedaRapida.nPredeterm = 1

 
Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                              "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                              "Cliente", 2, "Descripcion", 5500, 0, 0, "")
 
frmBusquedaRapida.Show vbModal
If wEnter = True And sCodigo <> "" Then
    LblCliente.Caption = sDescrip
    txtCodCliente.Text = sCodigo
Else
    LblCliente.Caption = ""
    txtCodCliente.Text = ""
End If
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, FrmTarjetaAproximidad.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, FrmTarjetaAproximidad.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, FrmTarjetaAproximidad.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, FrmTarjetaAproximidad.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, FrmTarjetaAproximidad.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, FrmTarjetaAproximidad.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & FrmTarjetaAproximidad.RsCabecera.AbsolutePosition & " de " & FrmTarjetaAproximidad.RsCabecera.RecordCount
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
Select Case Index
          Case Is = 0 ' Agregar
               Sw = True
               ActivarBotones (False)
               Blanquear Me
               'chkActivo.Value = 1
               txtMonto.Text = "0.00"
               'Cambia el Nombre del Primer Text
               txtCodigo.Enabled = True
                txtCodigo.SetFocus
          Case Is = 1 ' Grabar
               Dim nCorrela As String
               Dim validaCliente As String
               Dim xsaber As Integer
               
               'Chequea Datos
               If txtCodigo.Text = "" Then MsgBox "Ingrese Codigo de Tarjeta", vbExclamation, sMensaje: txtCodigo.SetFocus: Exit Sub
               If txtDetallado.Text = "" Then MsgBox "Ingrese Nombre del Tarjeta", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtCodCliente.Text = "" Then MsgBox "Seleccine el Cliente Asociado a la Tarjeta", vbExclamation, sMensaje: BtnBuscar.SetFocus: Exit Sub
               If CbEstado.Text = "" Then MsgBox "Seleccine estado de la Tarjeta", vbExclamation, sMensaje: CbEstado.SetFocus: Exit Sub
               
               If Sw Then
                  'Verifica si existe tarjeta
                  
                  nCorrela = Calcular("select (CodidoRFID) as Codigo from TTARJETASRFID where CodidoRFID = '" & txtCodigo.Text & "'", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Or nCorrela = "0" Then
                  Else
                    MsgBox "Codigo de Tarjeta ya existe", vbExclamation, sMensaje
                     Exit Sub
                  End If
                                    
                  validaCliente = Calcular("select (tCodigoDelivery) as Codigo from TDELIVERY where tCodigoDelivery = '" & txtCodCliente.Text & "'", Cn)
                                
                  If IsNull(validaCliente) Or validaCliente = "" Or validaCliente = "0" Then
                    MsgBox "Codigo de Cliente no existe", vbExclamation, sMensaje
                     Exit Sub
                  End If
                                    
                 
                    xsaber = 0
                    If CbEstado.Text = "Asignada" Then
                            xsaber = xsaber + 1
                    ElseIf CbEstado.Text = "Bloqueado" Then
                           xsaber = xsaber + 1
                    ElseIf CbEstado.Text = "Free" Then
                            xsaber = xsaber + 1
                        End If
                      
                    If xsaber = 0 Then
                        MsgBox "Debe Seleccionar un Estado, no editarlo", vbExclamation, sMensaje
                         Exit Sub
                    End If
                  
                  Sw = False
                                                       
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTARJETASRFID", "TARJETA DE PROXIMIDAD", "01", sUsuario, sPasa, "", _
                   "TTARJETASRFID", "TTARJETASRFID", "TARJETA DE PROXIMIDAD", "CodidoRFID", "Codigo Tarjeta", txtCodigo.Text, "Descripcion", "Descripcion Tarjeta", txtDetallado.Text, "MontoDisponible", "Monto", txtMonto.Text, "CodigoCliente", "Codigo Cliente", txtCodCliente.Text, "Estado", "Estado", CbEstado.Text)

                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If
                
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                   
                   
                   nMonto = txtMonto.Text
                   
                  'Cambiar el SQL
                  Isql = "insert into TTARJETASRFID( " & _
                         "CodidoRFID, fRegistro, UsuarioRegistro, FechaModificion, UsuarioModificion, Estado, MontoDisponible, Descripcion, CodigoCliente) " & _
                         "values (" & _
                                " '" & txtCodigo.Text & "', " & _
                                " getdate(), " & _
                                " '" & sUsuario & "', " & _
                                " getdate(), " & _
                                " '" & sUsuario & "', " & _
                                " '" & CbEstado.Text & "', " & _
                                " " & nMonto & ", " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtCodCliente.Text & "' " & ") "
           
                      Cn.Execute Isql
                      FrmTarjetaAproximidad.RsCabecera.Sort = "fRegistro ASC"
                      FrmTarjetaAproximidad.RsCabecera.Requery
                      FrmTarjetaAproximidad.RsCabecera.MoveLast
                      MsgBox "Registro Guardado", vbInformation, sMensaje
                      ActivarBotones (True)
                      cmdTexto.Caption = "Registro " & IIf(FrmTarjetaAproximidad.RsCabecera.RecordCount = 0, 0, FrmTarjetaAproximidad.RsCabecera.AbsolutePosition) & " de " & FrmTarjetaAproximidad.RsCabecera.RecordCount
              Else
              
              
               
                  validaCliente = Calcular("select (tCodigoDelivery) as Codigo from TDELIVERY where tCodigoDelivery = '" & txtCodCliente.Text & "'", Cn)
                                
                  If IsNull(validaCliente) Or validaCliente = "" Or validaCliente = "0" Then
                    MsgBox "Codigo de Cliente no existe", vbExclamation, sMensaje
                     Exit Sub
                  End If
                  
                  
                
                xsaber = 0
                
                If CbEstado.Text = "Asignada" Then
                        xsaber = xsaber + 1
                ElseIf CbEstado.Text = "Bloqueado" Then
                       xsaber = xsaber + 1
                ElseIf CbEstado.Text = "Free" Then
                        xsaber = xsaber + 1
                    End If
                  
                If xsaber = 0 Then
                    MsgBox "Debe Seleccionar un Estado, no editarlo", vbExclamation, sMensaje
                     Exit Sub
                End If
                
                  
               sPasa = txtCodigo.Text
                'Inserta Movimiento auditoria
                
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTARJETASRFID", "TARJETA DE PROXIMIDAD", "02", sUsuario, sPasa, "", _
                   "ID", "ID Tarjeta", LBLID.Caption, "CodidoRFID", "Codigo Tarjeta", txtCodigo.Text, "Descripcion", "Descripcion Tarjeta", txtDetallado.Text, "MontoDisponible", "Monto", txtMonto.Text, "CodigoCliente", "Codigo Cliente", txtCodCliente.Text, "Estado", "Estado", CbEstado.Text)

                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                        Exit Sub
                End If

                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
               
               
                 'Cambiar el SQL
                 Isql = "update TTARJETASRFID set " & _
                        "Descripcion ='" & txtDetallado.Text & "', " & _
                        "MontoDisponible =" & txtMonto.Text & ", " & _
                        "CodigoCliente ='" & txtCodCliente.Text & "', " & _
                        "Estado ='" & CbEstado.Text & "' ," & _
                        "UsuarioModificion = '" & sUsuario & "' ," & _
                        "FechaModificion = getdate() " & _
                        " where CodidoRFID = '" & txtCodigo.Text & "'"

                  Cn.Execute Isql
                  nPos = FrmTarjetaAproximidad.RsCabecera.Bookmark
                  FrmTarjetaAproximidad.RsCabecera.Requery
                  If FrmTarjetaAproximidad.RsCabecera.RecordCount = 0 Then
                     FrmTarjetaAproximidad.RsCabecera.Filter = adFilterNone
                  End If
                  FrmTarjetaAproximidad.RsCabecera.Bookmark = nPos
                  Screen.MousePointer = vbDefault
                  MsgBox "Registro Modificado", vbInformation, sMensaje
             End If
            
          
          Case Is = 3 ' Salir
               Unload Me
   End Select
End Sub
Sub ActivarBotones(ByVal Activa As Boolean)
    cmdNavegar(0).Enabled = Activa
    cmdNavegar(1).Enabled = Activa
    cmdNavegar(2).Enabled = Activa
    cmdNavegar(3).Enabled = Activa
    cmdNavegar(4).Enabled = Activa
    cmdNavegar(5).Enabled = Activa
    cmdOpcion(0).Enabled = Activa
'    cmdOpcion(2).Enabled = Activa
End Sub
Private Sub Form_Load()
 Screen.MousePointer = vbHourglass
    Centrar Me
    'Ingrese el SubTitulo
    Me.Caption = " Mantenimiento de Tarjeta de Proximidad "
    fraDetalle.Caption = Me.Caption
       
    Me.CbEstado.AddItem "Free"
    Me.CbEstado.AddItem "Asignada"
    Me.CbEstado.AddItem "Bloqueado"
       
    'Ingresar la Tabla
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       txtMonto.Text = "0.00"
       'chkActivo.Value = 1
       
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       txtCodigo.Enabled = False
       Asignar
    End If
    
    cmdTexto.Caption = "Registro " & FrmTarjetaAproximidad.RsCabecera.AbsolutePosition & " de " & FrmTarjetaAproximidad.RsCabecera.RecordCount
    Screen.MousePointer = vbDefault
End Sub
Sub Asignar()
    With FrmTarjetaAproximidad.RsCabecera
        'Cuadro de Texto
        LBLID.Caption = IIf(IsNull(!id), "", !id)
        txtCodigo.Text = IIf(IsNull(!CodidoRFID), "", !CodidoRFID)
        txtDetallado.Text = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtMonto.Text = Format(IIf(IsNull(!MontoDisponible), "0", !MontoDisponible), "##0.00")
        txtCodCliente.Text = IIf(IsNull(!codigoCliente), "", !codigoCliente)
        LblCliente.Caption = IIf(IsNull(!Cliente), "", !Cliente)
        CbEstado.Text = IIf(IsNull(!Estado), "", !Estado)
     
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set FrmTarjetaAproximidadDetalle = Nothing
End Sub

