VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCajaDetalle2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuraciòn Adicionales de Caja"
   ClientHeight    =   9990
   ClientLeft      =   2010
   ClientTop       =   1890
   ClientWidth     =   11955
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000001&
   Icon            =   "frmCajaDetalle2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   11955
   Visible         =   0   'False
   Begin VB.Frame fraDetalle 
      Height          =   9210
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11895
      Begin VB.TextBox txtDetallado 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         MaxLength       =   25
         TabIndex        =   20
         Text            =   " "
         Top             =   360
         Width           =   5970
      End
      Begin VB.TextBox txtCodigo 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   360
         Width           =   1170
      End
      Begin TabDlg.SSTab tabOpcion 
         Height          =   8130
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   14340
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Despachador"
         TabPicture(0)   =   "frmCajaDetalle2.frx":0442
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame21"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Integraciones"
         TabPicture(1)   =   "frmCajaDetalle2.frx":045E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Inforest App"
         TabPicture(2)   =   "frmCajaDetalle2.frx":047A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame2"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Extras"
         TabPicture(3)   =   "frmCajaDetalle2.frx":0496
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame4"
         Tab(3).Control(1)=   "Frame3"
         Tab(3).ControlCount=   2
         Begin VB.Frame Frame4 
            Caption         =   "Modificaciones Caja Rapida"
            Height          =   1455
            Left            =   -74880
            TabIndex        =   29
            Top             =   2280
            Width           =   4095
            Begin VB.CheckBox chkActivaAgrupadorCajaRapida 
               Caption         =   "Activar Agrupación Predeterminada"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   30
               Top             =   360
               Width           =   3855
            End
            Begin MSDataListLib.DataCombo cboAgrupadorDefault 
               Height          =   315
               Left            =   120
               TabIndex        =   31
               Top             =   840
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Activaciones Adicionales Punto de Venta"
            Height          =   1575
            Left            =   -74880
            TabIndex        =   26
            Top             =   600
            Width           =   4095
            Begin VB.CheckBox chkAlertaConsola 
               Caption         =   "Activar Alerta de Pedidosde Consola"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   1080
               Width           =   2895
            End
            Begin VB.CheckBox chkActivaVoucher 
               Caption         =   "Activar Obligatoriedad de Nº Voucher Requerido (Formulario de Pagos)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   27
               Top             =   480
               Width           =   3855
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Activaciones Inforest App"
            Height          =   855
            Left            =   -74760
            TabIndex        =   24
            Top             =   600
            Width           =   3615
            Begin VB.CheckBox chkTurnoCompartido 
               Caption         =   "Activación Turno Compartido"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Left            =   120
               TabIndex        =   25
               Top             =   360
               Width           =   2415
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Terminal"
            Height          =   1455
            Left            =   -74760
            TabIndex        =   21
            Top             =   600
            Width           =   3495
            Begin VB.CheckBox checkActivarIntegracionIzipay 
               Caption         =   "Activar Integración PinPad Izipay"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   240
               TabIndex        =   23
               Top             =   840
               Width           =   2895
            End
            Begin VB.CheckBox checkActivarIntegracionNiubiz 
               Caption         =   "Activar Integracion Inalámbrico Niubiz"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   240
               TabIndex        =   22
               Top             =   480
               Width           =   3075
            End
         End
         Begin VB.Frame Frame21 
            Caption         =   "Ordenes de Despacho Rappi"
            Height          =   1335
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   4575
            Begin MSDataListLib.DataCombo cboImpresoraDespachador 
               Height          =   315
               Left            =   1320
               TabIndex        =   18
               Top             =   420
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin VB.CheckBox chkListaOrdenes 
               Caption         =   "Activa lista de ordenes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   225
               Left            =   240
               TabIndex        =   17
               Top             =   960
               Width           =   2355
            End
            Begin VB.Label Label1 
               Caption         =   "Impresora :"
               Height          =   255
               Left            =   240
               TabIndex        =   19
               Top             =   480
               Width           =   1095
            End
         End
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Descripción :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   14
         Top             =   375
         Width           =   930
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código /"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   375
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   11895
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   9240
      Width           =   11955
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   6660
         TabIndex        =   3
         Top             =   60
         Width           =   6720
         Begin VB.CommandButton cmdNavegar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   3
            Left            =   5220
            Picture         =   "frmCajaDetalle2.frx":04B2
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   4
            Left            =   5700
            Picture         =   "frmCajaDetalle2.frx":09F4
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   5
            Left            =   6180
            Picture         =   "frmCajaDetalle2.frx":0F36
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmCajaDetalle2.frx":1478
            Style           =   1  'Graphical
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmCajaDetalle2.frx":19BA
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmCajaDetalle2.frx":1EFC
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            Caption         =   "Registro 0 de 0"
            Height          =   195
            Left            =   1590
            TabIndex        =   10
            Top             =   180
            Width           =   3495
         End
      End
      Begin VB.CommandButton cmdopcion 
         Caption         =   "Grabar"
         Height          =   615
         Index           =   0
         Left            =   9225
         Picture         =   "frmCajaDetalle2.frx":243E
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   60
         Width           =   1170
      End
      Begin VB.CommandButton cmdopcion 
         Caption         =   "Salir"
         Height          =   615
         Index           =   1
         Left            =   10520
         Picture         =   "frmCajaDetalle2.frx":2970
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmCajaDetalle2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsDespachador As Recordset
Dim RsAgrupadorCajaRapida As Recordset


Public CodigoCaja, DesCaja As String

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmCaja.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmCaja.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmCaja.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmCaja.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmCaja.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmCaja.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmCaja.RsCabecera.RecordCount = 0, 0, frmCaja.RsCabecera.AbsolutePosition) & " de " & frmCaja.RsCabecera.RecordCount
End Sub

Private Sub Form_Load()
    Asignar
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
    
   Select Case Index
         
          Case Is = 0 ' Grabar
        
                If CodigoCaja <> "" Then
                
                    Dim codCaja As Integer

                    codCaja = Calcular("Select count(*) as Codigo From TCAJA Where tcaja = '" & CodigoCaja & "'", Cn)

                    If codCaja <> 0 Then
                    
                        
                    
                        sPasa = CodigoCaja
                    
                        'Inserta Movimiento auditoria
                        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TCAJA", "CAJA", "01", sUsuario, sPasa, "", _
                                     "tDespachador", "Impresora Precuenta", cboImpresoraDespachador.BoundText, _
                                     "lOrdenesRappi", "Activa Lista Ordenes Rappi en Despachador", IIf(Me.chkListaOrdenes, "Verdadero", "Falso"))
                    
                        
                        'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                        If lAuditoria = False Then
                          Screen.MousePointer = vbDefault
                          Exit Sub
                        End If
                         
                        'Cambiar el SQL
                        'Isql = "update TCAJA set " & _
                        '      "tDespachador ='" & cboImpresoraDespachador.BoundText & "', lOrdenesRappi=" & Me.chkListaOrdenes.Value & "  where tCaja = '" & CodigoCaja & "'"
                                                     
                        'CSHM 01/09/2025 Activacion Niubiz o Izipay
                        Isql = "update TCAJA set " & _
                              "tDespachador ='" & cboImpresoraDespachador.BoundText & "', lOrdenesRappi=" & Me.chkListaOrdenes.Value & ",lActivaIntegracionNiubiz = " & checkActivarIntegracionNiubiz.Value & ",lActivaIntegracionIzipay = " & checkActivarIntegracionIzipay.Value & ",lTurnoCompartido = " & chkTurnoCompartido.Value & ", lActivaVoucher = " & chkActivaVoucher.Value & " ,lAlertaConsola = " & chkAlertaConsola.Value & " where tCaja = '" & CodigoCaja & "'"
                        
                        Cn.Execute Isql
                        
                        'Agrupador Caja Rapida
                        Isql = "update TCAJA set " & _
                              "lActivaAgrupadorCajaRapida ='" & Me.chkActivaAgrupadorCajaRapida.Value & "',tCodigoAgrupadorCajaRapida='" & cboAgrupadorDefault.BoundText & "' where tCaja = '" & CodigoCaja & "'"
                        
                        Cn.Execute Isql
                        
                        
                        'Inserta Movimiento auditoria Activacion Pinpad CSHM 01/09/2025
                        lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TCAJA", "CAJA", "01", sUsuario, sPasa, "", _
                                     "tIntegrador", "Integracion PinPad", cboImpresoraDespachador.BoundText, _
                                     "lOrdenesRappi", "Activa Lista Ordenes Rappi en Despachador", IIf(Me.chkListaOrdenes, "Verdadero", "Falso"), _
                                     "lActivaIntegracionNiubiz", "Activa Integracion Niubiz", IIf(Me.checkActivarIntegracionNiubiz, "Verdadero", "Falso"), _
                                     "lActivaIntegracionIzipay", "Activa Integracion Izipay", IIf(Me.checkActivarIntegracionIzipay, "Verdadero", "Falso"), _
                                     "lTurnoCompartido", "Activa Turno Compartido", IIf(Me.chkTurnoCompartido, "Verdadero", "Falso"), "lActivaVoucher", "Activa Obligatoriedad de Voucher Formulario de Pago", IIf(Me.chkActivaVoucher, "Verdadero", "Falso"), IIf(Me.chkAlertaConsola, "Verdadero", "Falso"))
                        
                        nPos = frmCaja.RsCabecera.Bookmark
                        frmCaja.RsCabecera.Requery
                        If frmCaja.RsCabecera.RecordCount = 0 Then
                           frmCaja.RsCabecera.Filter = adFilterNone
                        End If
                        frmCaja.RsCabecera.Bookmark = nPos
                        Screen.MousePointer = vbDefault
                        MsgBox "Registro Modificado", vbInformation, sMensaje
                    Else
                        MsgBox "No tiene una caja asignada, recuerda que primero debes registrar en mantenimiento de cajas para realizar configuración adicionales.", vbInformation, sMensaje
                        Unload Me
                    End If
                Else
                    MsgBox "No tiene una caja asignada, recuerda que primero debes registrar en mantenimiento de cajas para realizar configuración adicionales.", vbInformation, sMensaje
                    Unload Me
                End If
                   
          Case Is = 1 ' Salir
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
    cmdOpcion(2).Enabled = Activa
End Sub

Sub Asignar()

    txtCodigo.Text = CodigoCaja
    txtDetallado.Text = DesCaja
    
    LlenaCombos
    LlenaAgrupadorCajaRapida
    
    SubAsignar
    

End Sub
Sub SubAsignar()
    With frmCaja.RsCabecera
       
        cboImpresoraDespachador.BoundText = IIf(IsNull(!tDespachador), "", Trim(!tDespachador))
        chkListaOrdenes = IIf(!lOrdenesRappi = True, 1, 0)
        
        checkActivarIntegracionNiubiz = IIf(!lActivaIntegracionNiubiz = True, 1, 0)
        checkActivarIntegracionIzipay = IIf(!lActivaIntegracionIzipay = True, 1, 0)
        chkTurnoCompartido = IIf(!lTurnoCompartido = True, 1, 0)
        chkActivaVoucher = IIf(!lActivaVoucher = True, 1, 0)
        chkAlertaConsola = IIf(!lAlertaConsola = True, 1, 0)
        
        chkActivaAgrupadorCajaRapida = IIf(!lActivaAgrupadorCajaRapida = True, 1, 0)
        cboAgrupadorDefault.BoundText = IIf(IsNull(!tCodigoAgrupadorCajaRapida), "", !tCodigoAgrupadorCajaRapida)
        
    End With
End Sub


Sub LlenaAgrupadorCajaRapida()
    With cboAgrupadorDefault
         Isql = "select TCODIGO as codigo,ISNULL(tDetallado,'') as tDetallado from TTABLA  where TTABLA = 'ETIQUETA' and lActivo = 1"
         
         Set RsAgrupadorCajaRapida = Lib.OpenRecordset(Isql, Cn)
         
         Set .RowSource = RsAgrupadorCajaRapida
             .DataField = "tDetallado"
             .ListField = "tDetallado"
             .BoundColumn = "codigo"
    End With
End Sub



Sub LlenaCombos()
    With cboImpresoraDespachador
         Isql = "Select * from TIMPRESORA where tCaja = '" & CodigoCaja & "' order by tImpresora"
         
         Set RsDespachador = Lib.OpenRecordset(Isql, Cn)
         Set .RowSource = RsDespachador
             .DataField = "tDescripcion"
             .ListField = "tDescripcion"
             .BoundColumn = "tImpresora"
    End With
End Sub

