VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTipoDocumentoDetalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5085
   ClientLeft      =   2520
   ClientTop       =   2640
   ClientWidth     =   10380
   Icon            =   "frmTipoDocumentoDetalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   10380
   Begin TabDlg.SSTab SSTab1 
      Height          =   4305
      Left            =   2160
      TabIndex        =   28
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7594
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos Principales"
      TabPicture(0)   =   "frmTipoDocumentoDetalle.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDetalle"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Otras Activaciones"
      TabPicture(1)   =   "frmTipoDocumentoDetalle.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   35
         Top             =   360
         Width           =   7995
         Begin VB.CheckBox chkValidaUbigeo 
            Alignment       =   1  'Right Justify
            Caption         =   "Validación de Registro de Ubigeo y Urbanizacion:"
            Height          =   435
            Left            =   300
            TabIndex        =   52
            Top             =   1680
            Width           =   2835
         End
         Begin VB.TextBox txtCodPlantilla 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6450
            MaxLength       =   10
            TabIndex        =   51
            Text            =   " "
            Top             =   860
            Width           =   1080
         End
         Begin VB.TextBox txtMontoMinimo 
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
            Left            =   2880
            TabIndex        =   49
            Top             =   2700
            Width           =   735
         End
         Begin VB.CheckBox chkImprimeObservacion 
            Alignment       =   1  'Right Justify
            Caption         =   "Imprimir Observación en documento:"
            Height          =   195
            Left            =   240
            TabIndex        =   47
            Top             =   240
            Width           =   2895
         End
         Begin VB.TextBox txtMontoMaximo 
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
            Left            =   2880
            TabIndex        =   44
            Top             =   3050
            Width           =   735
         End
         Begin VB.TextBox txtImpresionFE 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1605
            Left            =   4080
            MaxLength       =   300
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            Text            =   "frmTipoDocumentoDetalle.frx":047A
            Top             =   1560
            Width           =   3480
         End
         Begin VB.CheckBox chkNotaCredito 
            Alignment       =   1  'Right Justify
            Caption         =   "Documento Nota de Crédito :"
            Height          =   195
            Left            =   765
            TabIndex        =   41
            Top             =   1440
            Width           =   2370
         End
         Begin VB.CheckBox chkAgenteRetencion 
            Alignment       =   1  'Right Justify
            Caption         =   "Impresión Texto Agente Retención :"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            ToolTipText     =   "El texto se registra en Parámetros Generales"
            Top             =   1200
            Width           =   2895
         End
         Begin VB.CheckBox chkCanjeNotaCredito 
            Alignment       =   1  'Right Justify
            Caption         =   "Pemitir Canjear por Nota de Crédito :"
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   480
            Width           =   2895
         End
         Begin VB.CheckBox chkQR 
            Alignment       =   1  'Right Justify
            Caption         =   "Impresión Código QR:"
            Height          =   195
            Left            =   1260
            TabIndex        =   11
            ToolTipText     =   "Estructura Bolivia: NitEmisor|NumFactura|NumAutorizacion|FechaEmisión|Total|BaseCreditoFiscal|CodControl|NitComprador|Descuento"
            Top             =   2400
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.CheckBox chkValidaEmail 
            Alignment       =   1  'Right Justify
            Caption         =   "Validación de Registro de Correo :"
            Height          =   200
            Left            =   420
            TabIndex        =   10
            Top             =   960
            Width           =   2715
         End
         Begin VB.TextBox txtCodigoFacturacion 
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
            Left            =   6450
            MaxLength       =   2
            TabIndex        =   14
            Text            =   " "
            Top             =   510
            Width           =   1080
         End
         Begin VB.CheckBox chkFacturacion 
            Alignment       =   1  'Right Justify
            Caption         =   "Facturación Electrónica :"
            Height          =   195
            Left            =   4440
            TabIndex        =   13
            Top             =   240
            Width           =   2205
         End
         Begin VB.CheckBox chkTransporte 
            Alignment       =   1  'Right Justify
            Caption         =   "Documento de Transporte :"
            Height          =   195
            Left            =   840
            TabIndex        =   9
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Código Plantilla (Carvajal) :"
            Height          =   195
            Index           =   9
            Left            =   4380
            TabIndex        =   50
            Top             =   880
            Width           =   1875
         End
         Begin VB.Label Label1 
            Caption         =   "Monto Mínimo Emisión:"
            Height          =   255
            Left            =   1200
            TabIndex        =   48
            Top             =   2760
            Width           =   1695
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            Caption         =   "(0= Ilimitado)"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   8
            Left            =   1080
            TabIndex        =   46
            Top             =   3480
            Width           =   1260
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Monto Máximo Emisión:"
            Height          =   195
            Index           =   7
            Left            =   1200
            TabIndex        =   45
            Top             =   3120
            Width           =   1665
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Texto Impresión Comprobante :"
            Height          =   195
            Index           =   6
            Left            =   4080
            TabIndex        =   43
            Top             =   1320
            Width           =   2205
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Código Facturación :"
            Height          =   195
            Index           =   5
            Left            =   4780
            TabIndex        =   36
            Top             =   555
            Width           =   1455
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
         Height          =   3735
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   7995
         Begin VB.Frame Frame2 
            Caption         =   "Doc. de Identidad Relacionados"
            Height          =   1575
            Left            =   4680
            TabIndex        =   37
            Top             =   1320
            Width           =   3255
            Begin VB.CheckBox chkOtros 
               Caption         =   "Otros Documentos de Identidad"
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   1080
               Width           =   2655
            End
            Begin VB.CheckBox chkRuc 
               Caption         =   "Identificador Tributario (RUC)"
               Height          =   255
               Left            =   120
               TabIndex        =   39
               Top             =   360
               Width           =   3015
            End
            Begin VB.CheckBox chkValidaRuc 
               Alignment       =   1  'Right Justify
               Caption         =   "Valida RUC :"
               Height          =   195
               Left            =   960
               TabIndex        =   38
               Top             =   720
               Width           =   1575
            End
         End
         Begin VB.CheckBox chkCliente 
            Alignment       =   1  'Right Justify
            Caption         =   "Pedir cliente de acuerdo al monto :"
            Height          =   195
            Left            =   180
            TabIndex        =   4
            Top             =   1804
            Width           =   2775
         End
         Begin VB.CheckBox chkActivo 
            Alignment       =   1  'Right Justify
            Caption         =   "Activo :"
            Height          =   195
            Left            =   2100
            TabIndex        =   7
            Top             =   2760
            Width           =   855
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
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   300
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
            Left            =   2760
            MaxLength       =   1
            TabIndex        =   2
            Text            =   " "
            Top             =   1052
            Width           =   1200
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
            Left            =   2760
            MaxLength       =   50
            TabIndex        =   1
            Text            =   " "
            Top             =   676
            Width           =   5100
         End
         Begin VB.TextBox txtSunat 
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
            MaxLength       =   2
            TabIndex        =   3
            Text            =   " "
            Top             =   1440
            Width           =   1200
         End
         Begin VB.TextBox txtMonto 
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
            Left            =   2760
            TabIndex        =   5
            Text            =   " "
            Top             =   2090
            Width           =   1200
         End
         Begin VB.CheckBox chkRegistroVenta 
            Alignment       =   1  'Right Justify
            Caption         =   "Registro de Venta :"
            Height          =   195
            Left            =   1260
            TabIndex        =   6
            Top             =   2400
            Width           =   1695
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Código :"
            Height          =   195
            Index           =   0
            Left            =   2055
            TabIndex        =   34
            Top             =   345
            Width           =   585
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Prefijo :"
            Height          =   195
            Index           =   2
            Left            =   2115
            TabIndex        =   33
            Top             =   1095
            Width           =   525
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Descripción :"
            Height          =   195
            Index           =   1
            Left            =   1710
            TabIndex        =   32
            Top             =   720
            Width           =   930
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Código Documento :"
            Height          =   195
            Index           =   3
            Left            =   1185
            TabIndex        =   31
            Top             =   1470
            Width           =   1455
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Monto :"
            Height          =   195
            Index           =   4
            Left            =   2100
            TabIndex        =   30
            Top             =   2130
            Width           =   540
         End
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   10320
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4335
      Width           =   10380
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
         Picture         =   "frmTipoDocumentoDetalle.frx":047C
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
         Left            =   7680
         Picture         =   "frmTipoDocumentoDetalle.frx":056E
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
         Left            =   5340
         Picture         =   "frmTipoDocumentoDetalle.frx":0670
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "frmTipoDocumentoDetalle.frx":0BA2
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   60
         Width           =   1170
      End
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   5190
         TabIndex        =   20
         Top             =   60
         Width           =   5250
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmTipoDocumentoDetalle.frx":10D4
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
            Picture         =   "frmTipoDocumentoDetalle.frx":1616
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
            Picture         =   "frmTipoDocumentoDetalle.frx":1B58
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4710
            Picture         =   "frmTipoDocumentoDetalle.frx":209A
            Style           =   1  'Graphical
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   4230
            Picture         =   "frmTipoDocumentoDetalle.frx":25DC
            Style           =   1  'Graphical
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3750
            Picture         =   "frmTipoDocumentoDetalle.frx":2B1E
            Style           =   1  'Graphical
            TabIndex        =   21
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
            TabIndex        =   27
            Top             =   180
            Width           =   2265
         End
      End
   End
   Begin VB.Image Image 
      Height          =   4455
      Left            =   0
      Picture         =   "frmTipoDocumentoDetalle.frx":3060
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2100
   End
End
Attribute VB_Name = "frmTipoDocumentoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Asignar()
    With frmTipoDocumento.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!codigo), "", !codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        txtResumido = IIf(IsNull(!Prefijo), "", !Prefijo)
        txtSunat = IIf(IsNull(!Sunat), "", !Sunat)
        txtMonto = IIf(IsNull(!nMontoMaximo), 0, !Monto)
       
        txtMontoMaximo = IIf(IsNull(!nMontoMaximo), 0, !nMontoMaximo)
        txtMontoMinimo = IIf(IsNull(!nMontoMinimo), 0, !nMontoMinimo)
    
        'Check Box
        chkTransporte.Value = IIf(IsNull(!Transporte), 0, IIf(!Transporte = True, 1, 0))
        Me.chkCanjeNotaCredito.Value = IIf(IsNull(!Canjear), 0, IIf(!Canjear = True, 1, 0))
        chkRegistroVenta.Value = IIf(IsNull(!RegistroVenta), 0, IIf(!RegistroVenta = True, 1, 0))
        chkCliente.Value = IIf(IsNull(!Cliente), 0, IIf(!Cliente = True, 1, 0))
        chkActivo = IIf(!lActivo = True, 1, 0)
        chkQR = IIf(!lImprimeQR = True, 1, 0)
        
        'FACTURACION ELECTRONICA
        txtCodigoFacturacion = IIf(IsNull(!tCodigoExterno), "", !tCodigoExterno)
        chkFacturacion.Value = IIf(IsNull(!lFacturacionE), 0, IIf(!lFacturacionE = True, 1, 0))
        txtImpresionFE = IIf(IsNull(!tImpresionFE), "", !tImpresionFE)
        
        chkNotaCredito.Value = IIf(IsNull(!lNotaCredito), 0, IIf(!lNotaCredito = True, 1, 0))
                
        chkValidaEmail.Value = IIf(IsNull(!lValidaEmail), 0, IIf(!lValidaEmail = True, 1, 0))
        
        chkAgenteRetencion.Value = IIf(IsNull(!lImpresionRetencion), 0, IIf(!lImpresionRetencion = True, 1, 0))
        chkRuc.Value = IIf(IsNull(!lidentidadruc), 0, IIf(!lidentidadruc = True, 1, 0))
        chkOtros.Value = IIf(IsNull(!lidentidadotros), 0, IIf(!lidentidadotros = True, 1, 0))
        chkValidaRuc = IIf(!lValidaRuc = True, 1, 0)
        
        chkImprimeObservacion = IIf(!lImpObs = True, 1, 0)
        
        'FACTURACION ELECTRONICA CARVAJAL
        txtCodPlantilla.Text = IIf(IsNull(!CodigoPlantilla), "", !CodigoPlantilla)
        Me.chkValidaUbigeo.Value = IIf(IsNull(!lValidaUbigeo), 0, IIf(!lValidaUbigeo = True, 1, 0))
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

Private Sub chkRuc_Click()
        If chkRuc.Value = 0 Then
            chkValidaRuc.Value = 0
        End If
End Sub
Private Sub chkValidaRuc_Click()
    If chkValidaRuc.Value = 1 Then
            If chkRuc.Value = 0 Then
                chkRuc.Value = 1
            End If
        Else
            chkRuc.Value = 0
    End If
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmTipoDocumento.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmTipoDocumento.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmTipoDocumento.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmTipoDocumento.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmTipoDocumento.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmTipoDocumento.grdGrilla
    End Select
   Asignar
   cmdTexto.Caption = "Registro " & IIf(frmTipoDocumento.RsCabecera.RecordCount = 0, 0, frmTipoDocumento.RsCabecera.AbsolutePosition) & " de " & frmTipoDocumento.RsCabecera.RecordCount
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
               chkQR.Value = 0
               chkFacturacion.Value = 0
               chkAgenteRetencion.Value = 0
                    
          Case Is = 1 ' Grabar
               Dim nCorrela As String
                    
               'Chequea Datos
               If txtDetallado.Text = "" Then MsgBox "Ingrese la Descripción", vbExclamation, sMensaje: txtDetallado.SetFocus: Exit Sub
               If txtResumido.Text = "" Then MsgBox "Ingrese el Prefijo", vbExclamation, sMensaje: txtResumido.SetFocus: Exit Sub
                    
               If Sw Then
                  'Obtiene el Numero de Orden
                  nCorrela = Calcular("select max(tCodigoTipoDocumento) as Codigo from TTIPODOCUMENTO", Cn)
                  If IsNull(nCorrela) Or nCorrela = "" Then
                      txtCodigo.Text = "01"
                  Else
                      txtCodigo.Text = Lib.Correlativo(nCorrela, 2)
                  End If
                  Sw = False
                   
                 sPasa = txtCodigo.Text
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPODOCUMENTO", "TIPO DOCUMENTO", "01", sUsuario, sPasa, "", _
                "tCodigoTipoDocumento", "Código Tipo Documento", sPasa, "tDescripcion", "Descripcion", txtDetallado.Text, "tPrefijo", "Prefijo Doc", txtResumido.Text, "tCodigoSunat", "Codigo Sunat", txtSunat.Text, "lPideCliente", "Flag Pide Cliente", IIf(chkCliente = 1, "Verdadero", "Falso"), _
                "nMonto", "Monto Minimo", val(txtMonto.Text), "lTransPorte", "Flag Documento de Transporte", IIf(chkTransporte = 1, "Verdadero", "Falso"), "lRegistroVenta", "Flag Registro de Venta", IIf(chkRegistroVenta = 1, "Verdadero", "Falso"), "lCanjearNotaCredito", "Flag Canje Nota Credito", IIf(Me.chkCanjeNotaCredito = 1, "Verdadero", "Falso"), _
                "lActivo", "Activo", IIf(Me.chkActivo = 1, "Verdadero", "Falso"), "lValidaRuc", "ValidaRuc", IIf(Me.chkValidaRuc = 1, "Verdadero", "Falso"), "lImprimeQR", "Imprime QR", IIf(Me.chkQR = 1, "Verdadero", "Falso"), "lImpresionRetencion", "Impresion Texto Retencion", IIf(chkAgenteRetencion = 1, "Verdadero", "Falso"), _
                "lIdentidadRuc", "Identidad Ruc", IIf(Me.chkRuc = 1, "Verdadero", "Falso"), "lIdentidadOtros", "Identidad Otros", IIf(Me.chkOtros = 1, "Verdadero", "Falso"), "nMontoMaximo", "Monto Maximo", val(txtMontoMaximo.Text), "tCodigoPlantilla", "Codigo Platilla Carvajal", val(txtCodPlantilla.Text), "lValidaUbigeo", "Valida Ubigeo y urbanizacion", IIf(Me.chkValidaUbigeo = 1, "Verdadero", "Falso"))
                
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.
                
                'Cambiar el SQL
                  Isql = "insert into TTIPODOCUMENTO( " & _
                         "tCodigoTipoDocumento, tDescripcion, tPrefijo, lPideCliente, tCodigoSunat, nMonto, lTransporte, lRegistroVenta, lActivo,lCanjearNotaCredito, lValidaRuc,lImprimeQR, tCodigoFacturacion, lFacturacionE, lNotaCredito, lValidaEmail, tImpresionFE,limpresionretencion,lidentidadruc,lidentidadotros,lImpObs,nmontomaximo,nmontominimo, tCodigoPlantilla, lValidaUbigeo) " & _
                         "values  ('" & txtCodigo.Text & "', " & _
                                " '" & txtDetallado.Text & "', " & _
                                " '" & txtResumido.Text & "', " & _
                                       chkCliente.Value & ", " & _
                                " '" & txtSunat.Text & "', " & _
                                       val(txtMonto.Text) & ", " & _
                                       chkTransporte.Value & ", " & _
                                       chkRegistroVenta.Value & ", " & _
                                       chkActivo.Value & ",  " & Me.chkCanjeNotaCredito.Value & ", " & Me.chkValidaRuc.Value & "," & Me.chkQR.Value & ", '" & txtCodigoFacturacion.Text & "'," & Me.chkFacturacion.Value & "," & Me.chkNotaCredito.Value & "," & Me.chkValidaEmail.Value & ",'" & txtImpresionFE.Text & "'," & Me.chkAgenteRetencion.Value & "," & chkRuc.Value & "," & chkOtros.Value & "," & chkImprimeObservacion.Value & ", " & val(txtMontoMaximo.Text) & ", " & val(txtMontoMinimo.Text) & ", " & val(txtCodPlantilla.Text) & ", " & chkValidaUbigeo.Value & ")"
                  Cn.Execute Isql
                  
                  
                  
                  
                  frmTipoDocumento.RsCabecera.Sort = "Codigo ASC"
                  frmTipoDocumento.RsCabecera.Requery
                  frmTipoDocumento.RsCabecera.MoveLast
                  MsgBox "Registro Guardado", vbInformation, sMensaje
                  ActivarBotones (True)
                  cmdTexto.Caption = "Registro " & IIf(frmTipoDocumento.RsCabecera.RecordCount = 0, 0, frmTipoDocumento.RsCabecera.AbsolutePosition) & " de " & frmTipoDocumento.RsCabecera.RecordCount
               Else
               sPasa = txtCodigo.Text
                'Inserta Movimiento auditoria
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPODOCUMENTO", "TIPO DOCUMENTO", "02", sUsuario, sPasa, "", _
                "tCodigoTipoDocumento", "Código Tipo Documento", sPasa, "tDescripcion", "Descripcion", txtDetallado.Text, "tPrefijo", "Prefijo Doc", txtResumido.Text, "tCodigoSunat", "Codigo Sunat", txtSunat.Text, "lPideCliente", "Flag Pide Cliente", IIf(chkCliente = 1, "Verdadero", "Falso"), _
                "nMonto", "Monto Minimo", val(txtMonto.Text), "lTransPorte", "Flag Documento de Transporte", IIf(chkTransporte = 1, "Verdadero", "Falso"), "lRegistroVenta", "Flag Registro de Venta", IIf(chkRegistroVenta = 1, "Verdadero", "Falso"), "lCanjearNotaCredito", "Flag Canje Nota Credito", IIf(Me.chkCanjeNotaCredito = 1, "Verdadero", "Falso"), _
                "lActivo", "Activo", IIf(Me.chkActivo = 1, "Verdadero", "Falso"), "lValidaRuc", "ValidaRuc", IIf(Me.chkValidaRuc = 1, "Verdadero", "Falso"), "lImprimeQR", "Imprime QR", IIf(Me.chkQR = 1, "Verdadero", "Falso"), "lImpresionRetencion", "Impresion Texto Retencion", IIf(chkAgenteRetencion = 1, "Verdadero", "Falso"), _
                "lIdentidadRuc", "Identidad Ruc", IIf(Me.chkRuc = 1, "Verdadero", "Falso"), "lIdentidadOtros", "Identidad Otros", IIf(Me.chkOtros = 1, "Verdadero", "Falso"), "nMontomaximo", "Monto maximo", val(txtMontoMaximo.Text), "tCodigoPlantilla", "Codigo Platilla Carvajal", val(txtCodPlantilla.Text), "lValidaUbigeo", "Valida Ubigeo y urbanizacion", IIf(Me.chkValidaUbigeo = 1, "Verdadero", "Falso"))
              If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
               
                  'Cambiar el SQL
                  Isql = "update TTIPODOCUMENTO set " & _
                         "tDescripcion ='" & txtDetallado.Text & "', " & _
                         "tPrefijo ='" & txtResumido.Text & "', " & _
                         "lPideCliente =" & chkCliente.Value & ", " & _
                         "tCodigoSunat ='" & txtSunat.Text & "', " & _
                         "nMonto =" & val(txtMonto.Text) & ",nMontoMaximo =" & val(txtMontoMaximo.Text) & ", nMontoMinimo =" & val(txtMontoMinimo.Text) & ", " & _
                         "lRegistroVenta =" & chkRegistroVenta.Value & ", " & _
                         "lTransporte =" & chkTransporte.Value & ", " & _
                         "lCanjearNotaCredito =" & Me.chkCanjeNotaCredito & ", " & _
                         "lActivo =" & chkActivo.Value & "," & _
                         "lValidaRuc =" & chkValidaRuc.Value & "," & _
                         "lIdentidadRuc =" & chkRuc.Value & "," & _
                         "lIdentidadOtros =" & chkOtros.Value & "," & _
                         "tCodigoFacturacion ='" & txtCodigoFacturacion.Text & "', " & _
                         "tCodigoPlantilla ='" & txtCodPlantilla.Text & "', " & _
                         "tImpresionFE ='" & txtImpresionFE.Text & "', " & _
                         "lFacturacionE =" & Me.chkFacturacion & ", " & _
                         "lNotaCredito =" & chkNotaCredito.Value & "," & _
                         "lValidaEmail =" & chkValidaEmail.Value & "," & _
                         "lValidaUbigeo =" & Me.chkValidaUbigeo.Value & "," & _
                         "lImpresionRetencion =" & Me.chkAgenteRetencion.Value & "," & _
                         "lImpObs =" & Me.chkImprimeObservacion.Value & "," & _
                         "lIMprimeqr =" & chkQR.Value & _
                         " where tCodigoTipoDocumento = '" & txtCodigo & "'"
                   Cn.Execute Isql
                   nPos = frmTipoDocumento.RsCabecera.Bookmark
                   frmTipoDocumento.RsCabecera.Requery
                   If frmTipoDocumento.RsCabecera.RecordCount = 0 Then
                      frmTipoDocumento.RsCabecera.Filter = adFilterNone
                   End If
                   frmTipoDocumento.RsCabecera.Bookmark = nPos
                   Screen.MousePointer = vbDefault
                   MsgBox "Registro Modificado", vbInformation, sMensaje
               End If
          
          Case Is = 2 ' Eliminar
               If frmTipoDocumento.RsCabecera.RecordCount = 0 Then
                  Exit Sub
               End If
               'Cambia el MsgBox
               If MsgBox("Seguro de Eliminar el Tipo Documento " & txtCodigo & "?", vbQuestion + vbOKCancel, sMensaje) = vbCancel Then
                  Exit Sub
               End If
                sPasa = txtCodigo.Text
                'Inserta Movimiento en este el codigo del registro viene del "txtcodigo.text"
                lAuditoria = RegistraMovimientoAuditoria(tModuloSeg, sMDB, "TTIPODOCUMENTO", "TIPO DOCUMENTO", "03", sUsuario, txtCodigo.Text, "", "tCodigoTipoDocumento", "Código Tipo Documento", txtCodigo.Text, "tDescripcion", "Descripcion", txtDetallado.Text)
                If lAuditoria = False Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
                'La Funcion RegistraMovimientoAuditoria devuelve true si se ejecuto correctamente.

               'Cambia el Delete
               Cn.Execute "delete from TTIPODOCUMENTO where tCodigoTipoDocumento = '" & txtCodigo & "'"
               frmTipoDocumento.RsCabecera.Requery
               If frmTipoDocumento.RsCabecera.RecordCount <> 0 Then
                  frmTipoDocumento.RsCabecera.MoveLast
                  Asignar
                  cmdTexto.Caption = "Registro " & IIf(frmTipoDocumento.RsCabecera.RecordCount = 0, 0, frmTipoDocumento.RsCabecera.AbsolutePosition) & " de " & frmTipoDocumento.RsCabecera.RecordCount
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
    Me.Caption = " Mantenimiento de Tipos de Documento"
    fraDetalle.Caption = Me.Caption
    
   Dim tNombreRuc As String
   tNombreRuc = Calcular("select tresumido as codigo from vtipoidentidad where codigo='02'", Cn)
   chkRuc.Caption = "Identificador Tributario (" & UCase(tNombreRuc) & ")"
   chkValidaRuc.Caption = "Valida " & UCase(tNombreRuc) & " :"
    
    If lFECarbajal Then
        Label(9).Visible = True
        txtCodPlantilla.Visible = True
    Else
        Label(9).Visible = False
        txtCodPlantilla.Visible = False
    End If
        
    If Sw = True Then
       ActivarBotones (False)
       Blanquear Me
       chkActivo.Value = 1
       chkQR.Value = 0
    Else
       'Cambiar la Busqueda y Nombre del formulario Cabecera
       ActivarBotones (True)
       Asignar
    End If
    If pais = "001" Then
        chkQR.Visible = True
    ElseIf pais = "003" Then
        Frame2.Caption = "Proceso de Facturacion"
        chkValidaRuc.Visible = False
        chkOtros.Visible = False
        chkRuc.Caption = "Filtrar cliente (I:Responsable Inscrito)"
        chkQR.Visible = False
        chkValidaUbigeo.Visible = False
    Else
        chkQR.Visible = False
    End If
    cmdTexto.Caption = "Registro " & IIf(frmTipoDocumento.RsCabecera.RecordCount = 0, 0, frmTipoDocumento.RsCabecera.AbsolutePosition) & " de " & frmTipoDocumento.RsCabecera.RecordCount
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cambia el Nombre del Formulario
    Set frmTipoDocumentoDetalle = Nothing
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




Private Sub txtMontoMaximo_KeyPress(KeyAscii As Integer)
    TabNext KeyAscii
   Numerico KeyAscii, txtMontoMaximo
End Sub

Private Sub txtMontoMinimo_KeyPress(KeyAscii As Integer)
    TabNext KeyAscii
   Numerico KeyAscii, txtMontoMinimo
End Sub

