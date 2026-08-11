VERSION 5.00
Object = "{B52C1CDE-38E9-11D5-98EC-00C0F01C6C81}#1.0#0"; "ifepson.ocx"
Begin VB.Form FrmRecargarTarjetaDetalle 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Recargar Tarjeta"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   11475
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   11415
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5445
      Width           =   11475
      Begin VB.PictureBox PicNavegacion 
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   4920
         TabIndex        =   32
         Top             =   60
         Width           =   4980
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   3480
            Picture         =   "FrmRecargarTarjetaDetalle.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   3960
            Picture         =   "FrmRecargarTarjetaDetalle.frx":0542
            Style           =   1  'Graphical
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   4440
            Picture         =   "FrmRecargarTarjetaDetalle.frx":0A84
            Style           =   1  'Graphical
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "FrmRecargarTarjetaDetalle.frx":0FC6
            Style           =   1  'Graphical
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "FrmRecargarTarjetaDetalle.frx":1508
            Style           =   1  'Graphical
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "FrmRecargarTarjetaDetalle.frx":1A4A
            Style           =   1  'Graphical
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.Label cmdTexto 
            Alignment       =   2  'Center
            Caption         =   "Registro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1470
            TabIndex        =   39
            Top             =   150
            Width           =   1845
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
         Left            =   6270
         Picture         =   "FrmRecargarTarjetaDetalle.frx":1F8C
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Left            =   7440
         Picture         =   "FrmRecargarTarjetaDetalle.frx":24BE
         Style           =   1  'Graphical
         TabIndex        =   30
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
         Left            =   8620
         Picture         =   "FrmRecargarTarjetaDetalle.frx":25C0
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Left            =   5100
         Picture         =   "FrmRecargarTarjetaDetalle.frx":26B2
         Style           =   1  'Graphical
         TabIndex        =   28
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
      Height          =   5415
      Left            =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   9675
      Begin VB.Frame FrmTarjeta 
         Height          =   4005
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   5370
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Autorización"
            Height          =   590
            Index           =   5
            Left            =   4080
            Picture         =   "FrmRecargarTarjetaDetalle.frx":2BE4
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   3360
            Width           =   1170
         End
         Begin VB.CommandButton cmdOpcion 
            Caption         =   "Descripción"
            Height          =   590
            Index           =   4
            Left            =   4080
            Picture         =   "FrmRecargarTarjetaDetalle.frx":2CE6
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   2700
            Width           =   1170
         End
         Begin VB.TextBox txtDescripcion 
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
            Height          =   285
            Left            =   120
            MaxLength       =   100
            TabIndex        =   46
            Text            =   " "
            Top             =   2955
            Width           =   3900
         End
         Begin VB.TextBox txtAutorizacion 
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
            Height          =   285
            Left            =   120
            MaxLength       =   50
            TabIndex        =   45
            Text            =   " "
            Top             =   3600
            Width           =   3900
         End
         Begin VB.CommandButton cmdReferencia 
            Caption         =   "Referencia"
            Height          =   590
            Left            =   4080
            TabIndex        =   43
            Top             =   2055
            Width           =   1185
         End
         Begin VB.TextBox txtReferencia 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   120
            MaxLength       =   30
            TabIndex        =   42
            Top             =   2325
            Width           =   3900
         End
         Begin VB.TextBox txtMonto 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   2760
            MaxLength       =   15
            TabIndex        =   40
            Text            =   " 0.00"
            Top             =   1455
            Width           =   2475
         End
         Begin VB.TextBox txtCodigoTarjeta 
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Top             =   470
            Width           =   3900
         End
         Begin VB.CommandButton BtnBuscar 
            Caption         =   "Buscar"
            Height          =   590
            Left            =   4080
            TabIndex        =   22
            Top             =   200
            Width           =   1185
         End
         Begin VB.Label LblCliente 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "*** Cliente ***"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   840
            Width           =   5175
         End
         Begin VB.Label LblSaldo 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "00.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   120
            TabIndex        =   53
            Top             =   1455
            Width           =   2415
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Actual :"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   1245
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Ingrese Tarjeta:"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   200
            Width           =   1110
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Descripción :"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   48
            Top             =   2700
            Width           =   930
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Autorización :"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   47
            Top             =   3360
            Width           =   960
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Referencia :"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   44
            Top             =   2085
            Width           =   870
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Monto :"
            Height          =   195
            Index           =   2
            Left            =   2760
            TabIndex        =   41
            Top             =   1245
            Width           =   540
         End
      End
      Begin VB.TextBox txtFecha 
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   380
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   2655
      End
      Begin VB.Frame Frame1 
         Caption         =   " Monto "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   3300
         Left            =   5600
         TabIndex        =   7
         Top             =   580
         Width           =   3980
         Begin VB.CommandButton cmdCancelacion 
            Caption         =   "OTRAS"
            Height          =   700
            Index           =   8
            Left            =   2640
            TabIndex        =   19
            Top             =   2480
            Width           =   1250
         End
         Begin VB.CommandButton cmdCancelacion 
            Height          =   700
            Index           =   7
            Left            =   1360
            TabIndex        =   18
            Top             =   2480
            Width           =   1250
         End
         Begin VB.CommandButton cmdCancelacion 
            Height          =   700
            Index           =   6
            Left            =   90
            TabIndex        =   17
            Top             =   2480
            Width           =   1250
         End
         Begin VB.CommandButton cmdCancelacion 
            Height          =   700
            Index           =   5
            Left            =   2640
            TabIndex        =   16
            Top             =   1740
            Width           =   1250
         End
         Begin VB.CommandButton cmdCancelacion 
            Height          =   700
            Index           =   4
            Left            =   1360
            TabIndex        =   15
            Top             =   1740
            Width           =   1250
         End
         Begin VB.CommandButton cmdCancelacion 
            Height          =   700
            Index           =   3
            Left            =   90
            TabIndex        =   14
            Top             =   1740
            Width           =   1250
         End
         Begin VB.CommandButton cmdCancelacion 
            Height          =   700
            Index           =   2
            Left            =   2640
            TabIndex        =   13
            Top             =   1000
            Width           =   1250
         End
         Begin VB.CommandButton cmdCancelacion 
            Height          =   700
            Index           =   1
            Left            =   1360
            TabIndex        =   12
            Top             =   1000
            Width           =   1250
         End
         Begin VB.CommandButton cmdCancelacion 
            Caption         =   "Efectivo"
            Height          =   700
            Index           =   10
            Left            =   1360
            TabIndex        =   11
            Top             =   270
            Width           =   1250
         End
         Begin VB.CommandButton cmdCancelacion 
            Caption         =   "Efectivo"
            Height          =   700
            Index           =   9
            Left            =   90
            TabIndex        =   10
            Top             =   270
            Width           =   1250
         End
         Begin VB.CommandButton cmdCancelacion 
            Caption         =   "Cheque MN"
            Height          =   700
            Index           =   11
            Left            =   2640
            TabIndex        =   9
            Top             =   270
            Width           =   1250
         End
         Begin VB.CommandButton cmdCancelacion 
            Caption         =   "Cheque ME"
            Height          =   700
            Index           =   12
            Left            =   90
            TabIndex        =   8
            Top             =   1000
            Width           =   1250
         End
      End
      Begin VB.Frame frmDocumento 
         BackColor       =   &H00808000&
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   4200
         Visible         =   0   'False
         Width           =   5415
         Begin VB.CommandButton cmdTipoDocumento 
            Height          =   735
            Index           =   2
            Left            =   1680
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton cmdTipoDocumento 
            Height          =   735
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chkImpuesto1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808000&
            Caption         =   "Check1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3240
            TabIndex        =   4
            Top             =   120
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CheckBox chkImpuesto2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808000&
            Caption         =   "Check1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   3240
            TabIndex        =   3
            Top             =   400
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CheckBox chkImpuesto3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808000&
            Caption         =   "Check1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   3240
            TabIndex        =   2
            Top             =   690
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin VB.Image imageCab 
         Height          =   375
         Left            =   7800
         Top             =   4800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imagepIE 
         Height          =   375
         Left            =   8400
         Top             =   4800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imageHash 
         Height          =   615
         Left            =   8880
         Top             =   4680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Index           =   1
         Left            =   6240
         TabIndex        =   26
         Top             =   280
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Tarjeta de pago seleccionada:"
         Height          =   255
         Left            =   5595
         TabIndex        =   25
         Top             =   3945
         Width           =   3735
      End
      Begin VB.Label lblnomtarjeta 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5595
         TabIndex        =   24
         Top             =   4200
         Width           =   3975
      End
   End
   Begin EPSON_Impresora_Fiscal.PrinterFiscal PrinterFiscalArgentina 
      Left            =   8760
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Image Image1 
      Height          =   5520
      Left            =   0
      Picture         =   "FrmRecargarTarjetaDetalle.frx":2DE8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1845
   End
   Begin VB.Image Image 
      Height          =   5835
      Left            =   0
      Picture         =   "FrmRecargarTarjetaDetalle.frx":DF95
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1785
   End
End
Attribute VB_Name = "FrmRecargarTarjetaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer
Dim sEstado As String
Dim nMonto As Double
Dim RsTarjeta As Recordset
Dim sTipoPago As String
Dim sTarjeta As String
Dim sMoneda As String
Dim RsTipoDocumento As Recordset
Dim RsIngreso As Recordset
Public TipoOper As String
Dim oComando As clsComando

Dim sMontoTotalAnticipo As Double

Private Sub BtnBuscar_Click()
On Error GoTo fin
 
   'If IIf(chkCover.value, 1, 0) = 1 Then
       Isql = "select CodidoRFID codigo, CodigoCliente  , t2.tApellido + ' ' + t2.tNombre descripcion , Estado , isnull(t1.MontoDisponible,0) MontoDisponible     from TTARJETASRFID t1 inner join TDELIVERY t2 on t1.CodigoCliente = t2.tCodigoDelivery  "
     
       Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Codigo", 1, "Codigo", 1700, 2, 0, "", _
                                                    "Descripción", 2, "Descripcion", 4500, 0, 0, "", _
                                                    "Monto Disponible", 2, "MontoDisponible", 1000, 1, 0, "", _
                                                    "Estado", 2, "Estado", 1000, 2, 0, "")
       frmBusquedaRapida.grdGrilla.Caption = "Tarjetas "
       frmBusquedaRapida.xbermur = "SI"
       frmBusquedaRapida.Show vbModal
       Sw = True

       If sCodigo <> "" Then
            txtCodigoTarjeta.Text = sCodigo
            LblSaldo.Caption = Format(sMontoBermur, "###,###,###,##0.00")
            LblCliente.Caption = sDescrip
       End If

   ' End If

Exit Sub
fin:
MsgBox (error)
End Sub

Private Sub cmdCancelacion_Click(Index As Integer)
Select Case Index
        Case Is <> 8
            sTipo = ""
            frmNumPad.Show vbModal
            nMonto = IIf(wEnter, val(sDescrip), nMonto)
            txtMonto.Text = Format(nMonto, "###,##0.00")
            If Index < 9 Then
                lblnomtarjeta.Caption = cmdCancelacion(Index).Caption
            End If
    End Select

   Select Case Index
          
          Case Is = 1 ' Tarjeta 1
               sTipoPago = "02"
               sTarjeta = "01"
               sMoneda = "01"
               
          Case Is = 2 ' Tarjeta 2
               sTipoPago = "02"
               sTarjeta = "02"
               sMoneda = "01"
'               lblnomtarjeta.Caption = cmdCancelacion(2).Caption
          Case Is = 3 ' Tarjeta 3
               sTipoPago = "02"
               sTarjeta = "03"
               sMoneda = "01"
'               lblnomtarjeta.Caption = cmdCancelacion(3).Caption
          Case Is = 4 ' Tarjeta 4
               sTipoPago = "02"
               sTarjeta = "04"
               sMoneda = "01"
'               lblnomtarjeta.Caption = cmdCancelacion(4).Caption
          Case Is = 5 ' Tarjeta 5
               sTipoPago = "02"
               sTarjeta = "05"
               sMoneda = "01"
'               lblnomtarjeta.Caption = cmdCancelacion(5).Caption
          Case Is = 6 ' Tarjeta 6
               sTipoPago = "02"
               sTarjeta = "06"
               sMoneda = "01"
               'lblnomtarjeta.Caption = cmdCancelacion(6).Caption
          Case Is = 7 ' Tarjeta 7
               sTipoPago = "02"
               sTarjeta = "07"
               sMoneda = "01"
               'lblnomtarjeta.Caption = cmdCancelacion(7).Caption
          Case Is = 8 ' Tarjeta 8
'               sTipoPago = "02"
'               sTarjeta = "08"
'               sMoneda = "01"

               
               Isql = "select tcodigotarjeta codigo, tdetallado,tresumido as descripcion,isnull(nfactorretencion,0) nfactorretencion, isnull(lpinpad,0) lpinpad, case when isnull(nboton,0)=0 then 99 else nboton end nboton, lactivo From TTARJETACREDITO Where lactivo = 1 order by tdetallado"
               Call ConfGrilla(2, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                               "Descripcion", 2, "Descripcion", 7000, 0, 0, "")
               frmBusquedaRapida.nPredeterm = 1
               frmBusquedaRapida.Show vbModal
                If wEnter Then
                    sTipoPago = "02"
                    sTarjeta = sCodigo
                    sMoneda = "01"
                    lblnomtarjeta.Caption = sDescrip
                    sTipo = ""
                    frmNumPad.Show vbModal
                    nMonto = IIf(wEnter, val(sDescrip), nMonto)
                    txtMonto.Text = Format(nMonto, "###,##0.00")
                End If

               
          Case Is = 9 ' Efectivo MN
               sTipoPago = "01"
               sTarjeta = ""
               sMoneda = "01"
               
          Case Is = 10 ' Efectivo ME
               sTipoPago = "01"
               sTarjeta = ""
               sMoneda = "02"
               
          Case Is = 11 ' Cheque MN
               sTipoPago = "03"
               sTarjeta = ""
               sMoneda = "01"
               
          Case Is = 12 ' Cheque ME
               sTipoPago = "03"
               sTarjeta = ""
               sMoneda = "02"
   End Select
End Sub

Private Sub CmdOpcion_Click(Index As Integer)
 Select Case Index
          Case Is = 3 'Salir
               Unload Me
          Case Is = 4 'KB Descripcion
               frmKeyBoard.txtResultado.Text = txtDescripcion.Text
               frmKeyBoard.Show vbModal
               txtDescripcion.Text = IIf(wEnter, sDescrip, txtDescripcion.Text)
               
          Case Is = 5 'KB Autorizacion
               frmKeyBoard.txtResultado.Text = txtAutorizacion.Text
               frmKeyBoard.Show vbModal
               txtAutorizacion.Text = IIf(wEnter, sDescrip, txtAutorizacion.Text)
   End Select
End Sub

Private Sub cmdReferencia_Click()
   sTipo = "Numero"
   frmNumPad.Show vbModal
   txtReferencia.Text = IIf(wEnter, sDescrip, txtReferencia.Text)
End Sub

Private Sub cmdTipoDocumento_Click(Index As Integer)
On Error GoTo fin
    Dim xUltimoCorrelativo As String
    nTotalPR = txtMonto.Text
    Dim imgcodigobarra As Image
    lDocumEquivaPrecuenta = False
    
    Dim RscadenaCodigoHash As Recordset
    
    Dim fso1 As Object
    
    Dim sSerie As String
    Dim sCorrela As String
    Dim sPrefijo As String
    Dim RsSuma As Recordset
    Dim sTipoDocumento As String
    Dim sImp As String
    Dim wConsumo As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim X As Integer
    Dim nRespuesta As Integer
    Dim nTotal As Double
    Dim sClienteFactura As String
    
    'sUsuarioAutoriza = sUsuario
    
    'Ecuador
    Dim sAutorizacion As String
    Dim fInicio As Date
    Dim fCaducidad As Date
   
    wConsumo = False

    sVar1 = sTipoPedido
    'txtBarra.SetFocus
    Dim msgError As String
    '--------------------------------------------------------------------------------
    Dim nMontoPedidoFacturarMInimo As Long
    Dim nMontoPedidoFacturar As Long
    Dim TimpresionDolaresDelivery As Boolean
    Dim sTipoDocum As String
    Dim lValidaEmail As Boolean
    Dim sEmail As String
    Dim lImprimeAlternativa  As Boolean
    Dim RsImpresion As Recordset
    Dim rstFuente As Recordset
    Dim numeroSerieImpresora As String
    Dim codigoImpresora As String
    Dim cadenaCodigoHash As String
    Dim sDetalleConsumo As String
    lImprimeAlternativa = False
    sDetalleConsumo = ""
        
        
   sMontoTotalAnticipo = CDbl(txtMonto.Text)
    'Chequea Consistencia
    RsTipoDocumento.Requery
    RsTipoDocumento.MoveFirst
    RsTipoDocumento.Find ("Descripcion='" & cmdTipoDocumento(Index).Caption & "'")
    If RsTipoDocumento.EOF Then
       MsgBox "Error: Configure los Documentos", vbCritical, sMensaje
       Exit Sub
    Else
       xlTipoDocumento = Calcular("Select lValidaRuc As Codigo From TTIPODOCUMENTO Where tCodigoTipoDocumento = '" & RsTipoDocumento!TTipoEmision & "'", Cn)
    End If
     
        
        
'validacionMontoMinimo
    nMontoPedidoFacturarMInimo = 0
    nMontoPedidoFacturarMInimo = Calcular("select isnull(nMontoMinimo,0) as codigo   from vTipoDocumento where Codigo='" & RsTipoDocumento!TTipoEmision & "'", Cn)
    If nMontoPedidoFacturarMInimo > 0 Then
        If nMontoPedidoFacturarMInimo >= CDbl(txtMonto.Text) Then
            MsgBox "El Monto a Facturar no llega al Minimo Permitido al Tipo de Documento"
            Exit Sub
        End If
    End If
    
    'monto maximo de recarga
    If CDbl(txtMonto.Text) > lMontoMaximo Then
        MsgBox "El monto de recarga esta superando, el monto Maximo de Recarga : " & lMontoMaximo & ""
        Exit Sub
    End If
    
    
    
    
    
    'validacionMontoMaximo
    nMontoPedidoFacturar = 0
    nMontoPedidoFacturar = Calcular("select isnull(nMontoMaximo,0) as codigo   from vTipoDocumento where Codigo='" & RsTipoDocumento!TTipoEmision & "'", Cn)
    If nMontoPedidoFacturar > 0 Then
        If nMontoPedidoFacturar <= CDbl(txtMonto.Text) Then
            MsgBox "El Monto a Facturar supera al Máximo Permitido al Tipo de Documento"
            Exit Sub
        End If
    End If

    TimpresionDolaresDelivery = False
    '-------- impresion en dolares si esta activo el check en el cliente delivery.
    If Calcular("select isnull(lEmisionMonedaExtranjera,0) as codigo from tdelivery where tcodigodelivery='" & sCliente & "'", Cn) Then
        If MsgBox("¿Desea Imprimir en " & sMonedaE & "?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            TimpresionDolaresDelivery = True
        Else
            TimpresionDolaresDelivery = False
        End If
    Else
        TimpresionDolaresDelivery = False
    End If

    lPagocortesiaAI = False

    If Not wConsumo And RsTipoDocumento!tFormulario <> "01" Then
    '------------- no es permitido emitir en otro formato
    Else
       '---------------------------------------------------------------------------
       '  ************************* EMISION POR TICKETERA *************************
       '---------------------------------------------------------------------------
         
       'Calcula el total de la cabecera
       nTotal = sMontoTotalAnticipo
        If RsTipoDocumento!Cliente And (RsTipoDocumento!Monto <= nTotal Or RsTipoDocumento!Monto = 0) Then
             'Consistencia Factura
            frmBusquedaRapida.nPredeterm = 1
            Isql = "exec usp_Inforest_ObtieneClientesFactura '" & sCliente & "','" & RsTipoDocumento!TTipoEmision & "', " & nTotal & ", '" & sCaja & "'"
        
            If Calcular("select COUNT(*) as codigo from TDELIVERYCLIENTE Where tCodigoDelivery='" & sCliente & "'", Cn) = 0 Then
               sTemp = ""
               frmBusquedaRapida.nPredeterm = 1
            End If
            
            Select Case pais 'ok
                    Case "001" 'Bolivia
                    
                        Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1200, 2, 0, "", _
                                                                         "Ident", 2, "tIdentidad", 1800, 2, 0, "", _
                                                                         "Cliente", 2, "Descripcion", 5200, 0, 0, "")
                    Case Else 'Peru, Ecuador
                        If lClub Then
                            Call ConfGrilla(4, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1000, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 4500, 0, 0, "", _
                                                                          "Enlace", 2, "tEnlace", 1100, 0, 0, "")
                        Else
                            Call ConfGrilla(3, frmBusquedaRapida.grdGrilla, "Código", 2, "Codigo", 1100, 2, 0, "", _
                                                                          "Ident", 2, "tIdentidad", 1600, 2, 0, "", _
                                                                          "Cliente", 2, "Descripcion", 5500, 0, 0, "")
                        End If
            End Select
             '------VALIDA CORREO----------
            sTipoDocum = RsTipoDocumento!TTipoEmision
            sTemp = ""
            frmBusquedaRapida.cmdOpcion(1).Enabled = True
            frmBusquedaRapida.cmdOpcion(2).Enabled = True
            frmBusquedaRapida.cmdOpcion(3).Enabled = True
            frmBusquedaRapida.Show vbModal
                
            If wEnter = True And sCodigo <> "" Then
                sTemp = sCliente
                sClienteFactura = sCodigo
                
                If pais = "000" Then
                    
                    'IMPRIME DNI
                    Dim RsTc1 As ADODB.Recordset
                    Set RsTc1 = New ADODB.Recordset
                    Set RsTc1 = Lib.OpenRecordset("usp_Inforest_ValidaClienteSel '" & sTipoDocum & "','" & sClienteFactura & "'", Cn)
                    If Not (RsTc1.EOF Or RsTc1.BOF) Then
                        RsTc1.MoveFirst
                        If RsTc1.Fields(0) <> "ok" Then
                            MsgBox "Error: El tipo de Identidad del Cliente no Corresponde al Tipo de Documento", vbCritical, sMensaje
                            Exit Sub
                        End If
                    End If
                End If
                lValidaEmail = Calcular("Select lValidaEmail As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn)
                
                If lValidaEmail = True Then
                     sEmail = Calcular("Select ISNULL(tcorreo,'') As codigo From vCLIENTE where Codigo ='" & sCodigo & "' ", Cn)
                  
                     If sEmail = "" Then
                        MsgBox "El cliente no tiene Email registrado", vbCritical, sMensaje
                        Exit Sub
                     End If
                     
                End If
                
                If Calcular("Select lValidaUbigeo As codigo From vTipoDocumento where Codigo='" & sTipoDocum & "'", Cn) = True Then
                    Dim TempUbigeo As String
                    Dim TempUrbaniza As String
                    TempUbigeo = Calcular("Select ISNULL(CodigoUbigeo,'') As codigo From vCLIENTE where Codigo ='" & sCodigo & "' ", Cn)
                    TempUrbaniza = Calcular("Select ISNULL(Urbanizacion,'') As codigo From vCLIENTE where Codigo ='" & sCodigo & "' ", Cn)
                    If Trim(TempUbigeo) = "" Or Trim(TempUrbaniza) = "" Then
                        MsgBox "El cliente no tiene Ubigeo ó Urbanizacion registrado, Favor de verificar!!!", vbCritical, sMensaje
                        Exit Sub
                    End If
                End If

                 Cn.Execute "usp_Inforest_ValidaDeliveryCliente '" & sCliente & "','" & sClienteFactura & "'"
            Else
                Exit Sub
            End If
        Else
            sClienteFactura = ""
        End If
    
        '****** INSERCCION DE PEDIDO DE ANTICIPO SEGUN CONFIGURACION ****************
        If InsertarPedidoAnticipo() = "" Then
            Exit Sub
        End If
        '***************************************************************************
        
        
        'Genera y Actualiza los Numero de Documento
        sSerie = RsTipoDocumento!tSerie
        sCorrela = Lib.Correlativo(RsTipoDocumento!tUltimoNumero, 9)
        sPrefijo = RsTipoDocumento!prefijo
        sTipoDocumento = RsTipoDocumento!TTipoEmision
        sImp = RsTipoDocumento!timpresora
        sDocumento = sPrefijo & sSerie & sCorrela
        sResumen = RsTipoDocumento!lResumen
              
        Set RsSuma = Lib.OpenRecordset("select sum(nVenta) as nVenta from DPEDIDO where (isnull(tFacturado,'0') = '0' or len(ltrim(tFacturado)) = 0 OR ISNULL(tFacturado,'')='') and tCodigoPedido ='" & sPedido & "' group by tCodigoPedido ", Cn)
        nCargo = Round(RsSuma!nVenta, 2)
    
    
        ''ELDCQ (INSERTA DOCUMENTO)
        If Not INSERTADOCUMENTO(sPedido, sDocumento, sTipoDocumento, IIf(sClienteFactura = "", "", sClienteFactura), "01", sCaja, sTurno, sSalon, sUsuario, sUsuario, Format(obtieneDiaContable, "yyyy/MM/dd"), "", sDetalleConsumo, IIf(TimpresionDolaresDelivery, 1, 0), "", "", "", Format(IIf(IsNull(RsTipoDocumento!fInicio), "01/01/1900", RsTipoDocumento!fInicio), "yyyy/mm/dd"), Format(IIf(IsNull(RsTipoDocumento!fCaducidad), "01/01/1900", RsTipoDocumento!fCaducidad), "yyyy/mm/dd"), tContribuyenteEspecial, 1, FechaServidor()) Then
            GoTo fin
        End If
    
        Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & sCorrela & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
        wEnter = True
        
'        Cn.Execute "Update MDOCUMENTO Set TRESERVA='" & frmReservaDetalle.txtCodigo.Text & "' where TDOCUMENTO='" & sDocumento & "'"
        xTipo = ""

        Isql = "insert into DPAGODOCUMENTO " & _
               "( tDocumento, tCorrelativo, tTurno, tTipoPago, tMoneda, tReferencia, nTipoCambio," & _
               " nMonto, npropina, tTarjeta, tNumero, tFechaVencimiento, fRegistro, tUsuario,fDiaContable,tNumeroATarjeta, ndolar ) " & _
               "Values(    '" & sDocumento & "',1 ,'" & sTurno & "', '" & sTipoPago & "', '" & sMoneda & "'," _
                        & "'', " _
                        & nTC & ", " & CDbl(txtMonto.Text) & ", 0, " _
                        & "'" & sTarjeta & "','" & Me.txtReferencia.Text & "','', getdate()," _
                        & "'" & sUsuario & "','" & Format(obtieneDiaContable, "yyyyMMdd") & "', ''," & IIf(sMoneda = "02", CDbl(txtMonto.Text), CDbl(txtMonto.Text)) & ")"
    
        Cn.Execute Isql
        Cn.Execute "Update MPEDIDO set fLlegada=getdate(),fEntrega=getdate(), tEstadoPedido = '02', lReplica = 1 where tCodigoPedido ='" & sPedido & "'"
        Cn.Execute "Update DPEDIDO set tFacturado = 'P', tDocumento = '" & sDocumento & "',tCortesia='' where tCodigoPedido ='" & sPedido & "' and (isnull(tFacturado,'0')='0' or len(ltrim(tFacturado))=0)"
        
                       
        Isql = "Update MDOCUMENTO set " & _
                  "tEstadoDocumento = '02', " & _
                  "nAbono = " & sMontoTotalAnticipo & ", " & _
                  "fPago = getdate()," & _
                  "nVuelto = 0 " & _
                  ",lreplica=1, TCORTESIA=''  where tDocumento ='" & sDocumento & "'"
        Cn.Execute Isql
        
        'PARA NO FISCALES
        Cn.Execute "UPDATE DPEDIDO SET lregistroventa=(select case when registroventa=0 then 0 else 1 end from vtipodocumento where codigo='" & sTipoDocumento & "') where tCodigoPedido ='" & sPedido & "' and  tDocumento ='" & sDocumento & "'"
        
         '-----------------------------------------------------------------------------------------------------
        If pais = "000" And lFEpape And IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
             If Not FacturarTCPIP(2, sDocumento, 0) Then
                 'Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                 Cn.Execute "delete MPEDIDO  where tCodigoPedido ='" & sPedido & "'"
                 Cn.Execute "delete DPEDIDO  where tCodigoPedido ='" & sPedido & "'"
                 Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                 Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                 
                 Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                 xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                 xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                 Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"

                GoTo fin
             End If
        End If
        '-------------------------------------------------------------------------------------------------------
        
         '============= nov 2010
        Dim imprimeDolaDocumentos As String
        imprimeDolaDocumentos = Calcular("select isnull(lequivadolares,0) as codigo from vtipodocumentoimpresora where tcaja='" & sCaja & "' and ttipoemision='" & sTipoDocumento & "' ", Cn)
        If imprimeDolaDocumentos = "Verdadero" Then
            lDocumEquivaPrecuenta = True
        Else
            lDocumEquivaPrecuenta = False
        End If
        
            If wConsumo = False And lDescripcionAlternativa = True Then
             If validaImpresionAlternativa(sDocumento) = False Then
                    If MsgBox("Desea imprimir descripción Alternativa? ", vbQuestion + vbYesNo + vbDefaultButton2, sMensaje) = vbYes Then
                          lImprimeAlternativa = True
                    End If
              End If
            End If
        '-------------------------------  SE INTRDUJO LAS CONSULTAS A  UN STORE PROCEDURE -------------------------
        
        ' ELDCQ 15/11/2017
        'Imprime Documentos
         If lImprimeAlternativa = False Then
            If lDocumentoAgrupado Then
                Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',1"
                'FACTURACION_E_PERU
                IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',2"
            Else
                Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',3"
                'FACTURACION_E_PERU
                IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',4"
            End If
        Else
            If lDocumentoAgrupado Then
                Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',5"
                'FACTURACION_E_PERU
                IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',6"
            Else
                Isql = "EXEC usp_Inforest_Impresion '" & sDocumento & "',7"
               'FACTURACION_E_PERU
                IsqlFact = "EXEC usp_Inforest_Impresion '" & sDocumento & "',8"
            End If
        End If
        
        
         '---------------------------------------
        Set RsImpresion = Lib.OpenRecordset(Isql, Cn)
        Set rstFuente = New ADODB.Recordset
        imageCab.Picture = Nothing
        
        
       If RsImpresion.RecordCount = 0 Then
          LimpiaRs
          MsgBox "No existen Datos a Imprimir", vbExclamation, sMensaje
       Else
        'SUNAT
        numeroSerieImpresora = obtieneNumeroSerieImpresora(sCaja, sImp)
        codigoImpresora = sImp
        'SUNAT
        Cn.Execute " update mdocumento set timpresora='" & codigoImpresora & "', tSerieImpresora='" & numeroSerieImpresora & "' where tdocumento ='" & sDocumento & "' "

               'FACTURACION_E_PERU
               If pais = "000" Then
                   If lFacturacionE Then
                   
                       If lFEOfisis Then 'OFISIS
                    
                       ElseIf lFESpring Then
                                                               
                       ElseIf lFECarbajal Then
                           
                       ElseIf lFEpape Then
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                If tCodigoFE = "000" Then
                                     If lQRFE Then
                                         Set imageHash.Picture = LoadPicture(CrearImagenQR(PapeTermico))
                                     Else
                                         If lImpresionCodigoBarras Then
                                             'Set imageHash.Picture = LoadPicture(ImagenQR(sDocumento))
                                         Else
                                             cadenaCodigoHash = PapeMatricial
                                         End If
                                     End If
                                 End If
                            End If
                       ElseIf lFEBiz Then
                            If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
                                 If Not INSERTA_FE_INFOREST(sDocumento, 1, DateTime.Date) Then '----CABECERA
                                     Cn.Execute "Update TMESA set tEstadoMesa = '02' where tCodigoMesa ='" & sMesa & "'"
                                     Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                     Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                     Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                                     xUltimoCorrelativo = Calcular("select MAX(tDocumento) as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                                     xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                                     Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                                     Cn.Execute "Delete from DPEDIDO  where tCodigoPedido ='" & sPedido & "' "
                                     Cn.Execute "Delete from MPEDIDO  where tCodigoPedido ='" & sPedido & "' "
                                     GoTo fin
                                  End If
                                  Sleep 1000
                                  'VALIDAR RESPUESTA DE CODIGO HASH Y CODIGO DE BARRA
                                  If tCodigoFE = "000" Then
                                      If lQRFE Then
                                          Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(3, sDocumento, 0))
                                      Else
                                          If lImpresionCodigoBarras Then
                                              Set imageHash.Picture = LoadPicture(QRHASH_FE_INFOREST(1, sDocumento, 0))
                                          Else
                                              cadenaCodigoHash = QRHASH_FE_INFOREST(2, sDocumento, 0)
                                          End If
                                      End If
                                  End If
                             End If
                            
                       Else 'INFOFACT

                       End If
                   End If
            End If
            '---------------------------------------
            'Configura la Impresora
            Imprimir (sImp)
            Printer.FontName = sFont
            Printer.FontBold = False
            
            If pais = "003" And Calcular("select isnull(tpuertoserial,'') as codigo from timpresora where tcaja='" & sCaja & "' and timpresora='" & sImp & "'", Cn) <> "" Then   ' Pais Argentina
                Dim RespDev As Integer
                RespDev = ImpresoraFiscalEpson(Me.PrinterFiscalArgentina, 2, RsImpresion)
                If RespDev = 0 Then
                    
                    Cn.Execute "Delete MDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    Cn.Execute "Delete DDOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    xUltimoCorrelativo = Calcular("select isnull(MAX(tDocumento),'000000000') as codigo from MDOCUMENTO where tcaja='" & sCaja & "' and tTipoDocumento='" & sTipoDocumento & "'", Cn)
                    xUltimoCorrelativo = Right(xUltimoCorrelativo, 9)
                    Cn.Execute "Update TTIPODOCUMENTOIMPRESORA Set tUltimoNumero = '" & xUltimoCorrelativo & "' where tTipoEmision ='" & sTipoDocumento & "' and tCaja ='" & sCaja & "'"
                    Cn.Execute "Delete from DPEDIDO  where tCodigoPedido ='" & sPedido & "' "
                    Cn.Execute "Delete from MPEDIDO  where tCodigoPedido ='" & sPedido & "' "
                    msgError = "Ocurrio un inconveniente en la facturacion con la impresora fiscal, " & vbNewLine & "revisar fuente de papel, conexion de impresora"
                    GoTo fin
                ElseIf RespDev = 1 Then
                    
                    Cn.Execute "Delete DPAGODOCUMENTO Where tDocumento= '" & sDocumento & "'"
                    Cn.Execute "Update MDOCUMENTO set tEstadoDocumento = '04', tUsuarioAnulado = '" & sUsuario & "', fRegistroAnulado = getdate() , tObservacion ='Anulacion por error en impresora fiscal', lreplica=1  where tDocumento ='" & sDocumento & "'"
                    Cn.Execute "Update MPEDIDO set tEstadoPedido = '03', tUsuarioAnulado = '" & sUsuario & "', fRegAnulado = getdate() , tObservacionAnulado ='Anulacion por error en impresora fiscal', tTurnoAnulado='" & sTurno & "', lreplica=1  where tCodigoPedido ='" & sPedido & "'"
                    Cn.Execute "update DPEDIDO set tfacturado='', tdocumento='' where tCodigoPedido ='" & sPedido & "' "
                    msgError = "Error en el proceso de impresion, " & vbNewLine & "Corte de energia u otro, revisar fuente de papel, conexion de impresora"
                    GoTo fin
                End If
            Else
            
            If wConsumo Then
               If RsTipoDocumento!tFormulario = "01" Then
               
                  If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                      'FACTURACION ELECTRONICA
                      If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                           ImprimeFacturaConsumoElectronico RsImpresion, sDetalleConsumo, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                      Else
                           ImprimeFacturaConsumoT RsImpresion, sDetalleConsumo, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                      End If
                     
                      NFactura = sCorrela
                      'lblFactura.Caption = NFactura
                  Else
                      'FACTURACION ELECTRONICA
                      If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                          ImprimeBoletaConsumoElectronico RsImpresion, sDetalleConsumo, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                      Else
                          ImprimeBoletaConsumoT RsImpresion, sDetalleConsumo, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                      End If
                  End If
                  
               Else
                               
                  
               End If
            Else
               If RsTipoDocumento!tFormulario = "01" Then
                  If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
            
                          'FACTURACION ELECTRONICA
                          If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                             ImprimeFacturaElectronica RsImpresion, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                          Else
                             ImprimeFacturaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                          End If
                     
                        NFactura = sCorrela
                        'lblFactura.Caption = NFactura
                          
                  ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                          If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                             ImprimeCortesia RsImpresion, "00", imageCab, imagepIE
                          End If
                  Else
            
                          'FACTURACION ELECTRONICA
                          If IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) And tCodigoFE <> "999" Then
                             ImprimeBoletaElectronica RsImpresion, imageHash, sTipoDocumento, imageCab, imagepIE, cadenaCodigoHash, TimpresionDolaresDelivery
                          Else
                             ImprimeBoletaT RsImpresion, sTipoDocumento, imageCab, imagepIE, TimpresionDolaresDelivery
                          End If
                        
                  End If
                  
               Else
                  If RsTipoDocumento!Cliente And RsTipoDocumento!Monto = 0 Then
                     ImprimeFacturaN RsImpresion, sEmpresa, sTipoDocumento
                     NFactura = sCorrela
                     'lblFactura.Caption = NFactura
                  ElseIf RsTipoDocumento!TTipoEmision = "00" Then
                     If MsgBox("Deseas imprimir el Voucher", vbQuestion + vbYesNo, sMensaje) = vbYes Then
                        ImprimeCortesia RsImpresion, "00", imageCab, imagepIE
                     End If
                  Else
                     ImprimeBoletaN RsImpresion, sEmpresa, sTipoDocumento
                  End If
                  
               End If
               
            End If
            
         End If ' fin de evaluacion de pasi argentina 003 - impresion fiscal
         
       End If
       
        LimpiaRs
        
    End If


    If sCliente <> "" Then
       Cn.Execute "Update TDELIVERY set tCodigoCliente = '" & sClienteFactura & "' where tCodigoDelivery ='" & sTemp & "'"
    End If
    sTemp = ""
    
    '-----------------------
     If pais = "000" And lFEpape And IIf(RsTipoDocumento!lFacturacionElectronica = True, 1, 0) Then
          If Not FacturarTCPIP(3, sDocumento, 0) Then
             MsgBox ("La confirmacion ha fallado favor de contactarse con paperlees"), vbInformation, sMensaje
          End If
     End If
     '------------------------
    
    
    If sDocumento <> "" Then
        Isql = "insert into TMOVIMIENTOTARJETASRFID " & _
           "( CodidoRFID, Tipo, DocReferencia, CodReferencia, CodDispensador, Cantidad, fRegistro," & _
           " MontoIngreso, MontoSalida, MontoAnterior, MontoFinal ) " & _
           " Values (  " & _
           " '" & txtCodigoTarjeta.Text & "', " & _
           " 'R' , " & _
           " '" & sDocumento & "' ,  1 , 0 , 1 , getdate() , " & _
           " '" & CDbl(txtMonto.Text) & "', " & _
           " " & 0 & ", " & _
           " '" & CDbl(Me.LblSaldo.Caption) & "', " & _
           " '" & CDbl(txtMonto.Text) + CDbl(Me.LblSaldo.Caption) & "')"
        Cn.Execute Isql
        
        
         Isql = "update TTARJETASRFID set " & _
                        " MontoDisponible =  MontoDisponible + " & CDbl(txtMonto.Text) & " " & _
                        " where CodidoRFID = '" & txtCodigoTarjeta.Text & "'"
        Cn.Execute Isql
        
        
    End If
    
    FrmRecargarTarjeta.RsCabecera.Requery
    If FrmRecargarTarjeta.RsCabecera.RecordCount = 0 Then
       FrmRecargarTarjeta.RsCabecera.Filter = adFilterNone
    End If
    
    Screen.MousePointer = vbDefault
      
    MsgBox "Recargar Realizada Correctamente!!", vbInformation, sMensaje
    Limpiar
    
'    Unload Me
    Exit Sub

fin:

    Screen.MousePointer = vbDefault
    Call Log_Inforest("PUNTO VENTA", "EMISION DE DOCUMENTO PUNTO DE VENTA", sPedido, "", sDocumento, Estado & " " & error & msgError, "", "FALLA AL GENERAR DOCUMENTO EN PUNTO DE VENTA ", sUsuario)
    MsgBox "Error: " & error & msgError & vbNewLine & "- Proceso: Emision de Documentos Punto Venta", vbCritical, sMensaje

End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
Centrar Me
   txtFecha.Text = FechaServidor()
 'Tarjetas de Crédito
   Isql = "select * from TTARJETACREDITO where nBoton>0 and lActivo = 1 Order by nBoton"
   Set RsTarjeta = Lib.OpenRecordset(Isql, Cn)
   Call AsignaTarjeta(7, RsTarjeta, cmdCancelacion())
   
   nMonto = 0
   cmdCancelacion(9).Caption = cmdCancelacion(9).Caption & " " & sMonN

   If sMonE = "" Then
      cmdCancelacion(10).Visible = False
   Else
      cmdCancelacion(10).Caption = cmdCancelacion(10).Caption & " " & sMonE
   End If
   
   LlenarTarjeta
   
   cmdTexto.Caption = "Registro " & FrmRecargarTarjeta.RsCabecera.AbsolutePosition & " de " & FrmRecargarTarjeta.RsCabecera.RecordCount
   Screen.MousePointer = vbDefault
End Sub
Private Sub LlenarTarjeta()
    Me.fraDetalle.Caption = ""
    Me.Caption = "Recargar Tarjeta"
'    Me.cmdOpcion(6).Visible = False
'    Me.Label(6).Visible = False
'    Me.Label(7).Visible = False
'    Me.Label(0).Visible = False
'    Me.Label(8).Visible = False
''    Me.cmdBuscar.Visible = False
'    Me.txtPedido.Visible = False
'    Me.txtRecibo.Visible = False
'    Me.txtTurno.Visible = False
    
    Me.Label(4).Caption = "Observacion:"
    Me.Label(5).Caption = "Glosa:"
    Me.cmdOpcion(5).Caption = "Observacion"
    Me.cmdOpcion(5).Caption = "Glosa"
    
    Set RsTipoDocumento = Lib.OpenRecordset("select top 2  * from vTipoDocumentoImpresora where tCaja ='" & sCaja & "' and Transporte=0 And lNotaCredito = 0 And lActivo = 1 and tTipoEmision<>'00' order by tTipoEmision", Cn)
    AsignaComando 2, RsTipoDocumento, cmdTipoDocumento()
    
    'cmdTipoDocumento(1).Visible = True
    'cmdTipoDocumento(2).Visible = True
    frmDocumento.Visible = True
    cmdOpcion(0).Visible = False
    cmdOpcion(1).Visible = False
    cmdOpcion(2).Visible = False
    If nTC = 0 Then
        cmdCancelacion(12).Enabled = False
    End If
    If sImpuesto1 <> "" Then
        Me.chkImpuesto1.Visible = True
        Me.chkImpuesto1.Caption = sImpuesto1
        Me.chkImpuesto1.value = 1
    End If
    If sImpuesto2 <> "" Then
        Me.chkImpuesto2.Visible = True
        Me.chkImpuesto2.Caption = sImpuesto2
        Me.chkImpuesto2.value = 0
    End If
    If sImpuesto3 <> "" Then
        Me.chkImpuesto3.Visible = True
        Me.chkImpuesto3.Caption = sImpuesto3
        Me.chkImpuesto3.value = 0
    End If

End Sub
Public Sub AsignaTarjeta(nBoton As Integer, RsAsigna As Recordset, cmdBoton As Object)
   Screen.MousePointer = vbHourglass
   Dim i As Integer
   With RsAsigna
        If .RecordCount > 0 Then
           For i = 1 To nBoton
               .MoveFirst
               .Find "nboton = " & Trim(str(i))
               If Not .EOF Then
                  cmdBoton(i).Visible = True
                  cmdBoton(i).Caption = !tResumido
                Else
                    cmdBoton(i).Visible = False
                End If
           Next i
        Else
           For i = 1 To nBoton
               cmdBoton(i).Visible = False
           Next i
        End If
  End With
  Screen.MousePointer = vbDefault
End Sub
Sub Limpiar()
    txtCodigoTarjeta.Text = ""
    txtMonto.Text = "0.00"
    Me.LblSaldo.Caption = "0.00"
    txtReferencia.Text = ""
    txtDescripcion.Text = ""
    txtAutorizacion.Text = ""
    lblnomtarjeta.Caption = ""
    LblCliente.Caption = "*** Cliente ***"
End Sub

Function InsertarPedidoAnticipo() As String
On Error GoTo fin
    Dim CodProd As String
    Dim RsProducto As Recordset
    Dim nPVenta As Double
    Dim nPBase As Double
    Dim nImpuesto1 As Double
    Dim nImpuesto2 As Double
    Dim nImpuesto3 As Double
    Dim nRecargo As Double
    Dim nDescuento As Double
    Dim nOficial As Double
    Dim nCantidad As Double
    Dim Acumulado As Double
    '**********
    Dim nValor As Double
    Dim lImp1 As Boolean
    Dim lImp2 As Boolean
    Dim lImp3 As Boolean
    Dim Oficial As Double

    '**********
     
    CodProd = Calcular("select top 1 isnull(tCodTarjeta,'') as codigo from tparametro", Cn)
     
    If CodProd = "" Then
        MsgBox "No hay producto configurado como Recarga de tarjeta en parametros generales!!!"
        Exit Function
    End If
    '****** RECARGA PRODUCTOS Y PRODUCTO ANTICIPO DE PARAMETRO GENERAL ******
    Set RsProducto = Lib.OpenRecordset("usp_Inforest_ObtieneProductos", Cn)
    RsProducto.Find "Codigo = '" & Trim(CodProd) & "'"
    
    If RsProducto.BOF Or RsProducto.EOF Then
        Exit Function
    End If

    'Cn.BeginTrans
    '****** GENERA EL PROCESO DE INSERCCION DE PEDIDO EN EL SISTEMA *****
    sPedido = ""
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("spIns_MPEDIDO", Cn) Then
       Set oComando = Nothing
       GoTo fin
    End If
    oComando.CreateParameter "@tCliente", adVarChar, adParamInput, 7, sCliente
    oComando.CreateParameter "@tTipoPedido", adVarChar, adParamInput, 2, "01"
    oComando.CreateParameter "@lPrioridad", adBoolean, adParamInput, 1, 0
    oComando.CreateParameter "@tTipoAtencion", adVarChar, adParamInput, 2, "01"
    oComando.CreateParameter "@tMesa", adVarChar, adParamInput, 3, ""
    oComando.CreateParameter "@tMozo", adVarChar, adParamInput, 4, ""
    oComando.CreateParameter "@tMotorizado", adVarChar, adParamInput, 4, ""
    oComando.CreateParameter "@tCaja", adVarChar, adParamInput, 3, sCaja
    oComando.CreateParameter "@tSalon", adVarChar, adParamInput, 2, sSalon
    oComando.CreateParameter "@tTurno", adVarChar, adParamInput, 10, sTurno
    oComando.CreateParameter "@tObservacion", adVarChar, adParamInput, 250, txtDescripcion.Text
    oComando.CreateParameter "@nTiempo", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, Right(sUsuario, 15)
    oComando.CreateParameter "@nAdulto", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@nNino", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@nMesa", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@tPuntoVenta", adVarChar, adParamInput, 2, ""
    oComando.CreateParameter "@tHabitacion", adVarChar, adParamInput, 6, ""
    oComando.CreateParameter "@tReserva", adVarChar, adParamInput, 6, ""
    oComando.CreateParameter "@tPasajero", adVarChar, adParamInput, 50, ""
    oComando.CreateParameter "@tCompania", adVarChar, adParamInput, 5, ""
    oComando.CreateParameter "@tContacto", adVarChar, adParamInput, 4, ""
    oComando.CreateParameter "@nDescuento", adDouble, adParamInput, 10, 0
    oComando.CreateParameter "@tDescuento", adVarChar, adParamInput, 3, ""
    oComando.CreateParameter "@tObservacionDescuento", adVarChar, adParamInput, 250, ""
    oComando.CreateParameter "@tAutorizaDescuento", adVarChar, adParamInput, 15, ""
    oComando.CreateParameter "@nTiempoDelivery", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@tTienda", adVarChar, adParamInput, 3, ""
    oComando.CreateParameter "@fDiaContable", adDate, adParamInput, 10, obtieneDiaContable
    oComando.CreateParameter "@fProgramacion", adDate, adParamInput, 20, Format(FechaServidor, "dd/MM/yyyy HH:mm")
    oComando.CreateParameter "@tCodigoInvitado", adVarChar, adParamInput, 10, ""
    oComando.CreateParameter "@tCodigopariente", adVarChar, adParamInput, 7, ""
    oComando.CreateParameter "@tEntregarA", adVarChar, adParamInput, 20, ""
    oComando.CreateParameter "@nTiempoAntesEnvio", adInteger, adParamInput, 10, 0
    oComando.CreateParameter "@nMontoMaximo", adVarChar, adParamInput, 250, 0
    oComando.CreateParameter "@tPedido", adVarChar, adParamOutput, 10, sPedido
    oComando.CreateParameter "@codigoOrigenVentas", adVarChar, adParamInput, 2, ""

    If Not oComando.GetParamOK Then
       Set oComando = Nothing
       GoTo fin
    End If
    If Not oComando.ExecSP Then
       Set oComando = Nothing
       GoTo fin
    Else
       sPedido = oComando.GetParameterValue("@tPedido")
    End If
                                                                             
    If sPedido <> "" Then
        nRecargo = 0
        nDescuento = 0
        
        lImp1 = Me.chkImpuesto1.value
        lImp2 = Me.chkImpuesto2.value
        lImp3 = Me.chkImpuesto3.value
      
        nValor = 0
        nValor = nValor + IIf(lImp1, nPorcentaje1, 0)
        nValor = nValor + IIf(lImp2, nPorcentaje2, 0)
        nValor = nValor + IIf(lImp3, nPorcentaje3, 0)


        If Me.chkImpuesto1.value = 1 Then
            lImp1 = True
        Else
            lImp1 = False
        End If
        
        If Me.chkImpuesto2.value = 1 Then
            lImp2 = True
        Else
            lImp2 = False
        End If
        
        If Me.chkImpuesto3.value = 1 Then
            lImp3 = True
        Else
            lImp3 = False
        End If

       nPVenta = sMontoTotalAnticipo
       nOficial = sMontoTotalAnticipo
       xCantidad = 1
       
       Oficial = IIf(IsNull(RsProducto!nprecioVenta), 0, RsProducto!nprecioVenta)
       nOficial = IIf(RsProducto!tMoneda = "02", nOficial * nTC, nOficial)
    
       Select Case pais 'ok
           Case "001" 'Bolivia
                   nValor = (nValor / 100)
                   nImpuesto1 = IIf(lImp1, nPVenta * nPorcentaje1 / 100, 0)
                   nImpuesto2 = IIf(lImp2, nPVenta * nPorcentaje2 / 100, 0)
                   nImpuesto3 = IIf(lImp3, nPVenta * nPorcentaje3 / 100, 0)
                   nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
           
           Case Else 'Peru, Ecuador
                   nValor = 1 + (nValor / 100)
                   nImpuesto1 = IIf(lImp1, nPVenta / nValor * nPorcentaje1 / 100, 0)
                   nImpuesto2 = IIf(lImp2, nPVenta / nValor * nPorcentaje2 / 100, 0)
                   nImpuesto3 = IIf(lImp3, nPVenta / nValor * nPorcentaje3 / 100, 0)
                   nPBase = nPVenta - nImpuesto1 - nImpuesto2 - nImpuesto3
       End Select
    
        Dim nInsumo As Double
        Dim nGasto As Double
        Dim nMObra As Double
    

       nInsumo = IIf(IsNull(RsProducto!nInsumo), 0, RsProducto!nInsumo)
       nGasto = IIf(IsNull(RsProducto!nGasto), 0, RsProducto!nGasto)
       nMObra = IIf(IsNull(RsProducto!nManoObra), 0, RsProducto!nManoObra)


        Isql = "insert into DPEDIDO " & _
           "(tCodigoPedido, tTipoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, " & _
           "nPrecioNeto, nRecargo, nDescuento, nPrecioOficial, nPrecioImpuesto1, nPrecioImpuesto2, nPrecioImpuesto3, nPrecioVenta, " & _
           "nCantidad, nVenta, nImpuesto1, nImpuesto2, nImpuesto3, " & _
           "lImprime, tArea, lImprimeArea, lCombinacion, nCombinacion, tEstadoItem, tComanda, fRegistro, tMozoD, tUsuarioD, nInsumo, nGasto, nManoObra, nOrden, tUnidadNegocio, tOferta, tsubalmacen,fdiacontable,tcajad , tobservacion) " & _
           "Values(   '" & sPedido & "', " _
                   & "'01', '001', '" & CodProd & "', " _
                   & "'" & IIf(IsNull(RsProducto!tgrupo), "", RsProducto!tgrupo) & "', " _
                   & "'" & IIf(IsNull(RsProducto!tSubGrupo), "", RsProducto!tSubGrupo) & "', " _
                   & nPBase & ", " & nRecargo & ", " & nDescuento & ", " & nOficial & ", " _
                   & nImpuesto1 & ", " & nImpuesto2 & ", " & nImpuesto3 & ", " _
                   & nPVenta & ", " & xCantidad & ", " & nPVenta * xCantidad & ", " _
                   & nImpuesto1 * xCantidad & ", " & nImpuesto2 * xCantidad & ", " & nImpuesto3 * xCantidad & ", " _
                   & "0, '" & RsProducto!tArea & "', " _
                   & IIf(RsProducto!lImprimeArea, -1, 0) & "," _
                   & IIf(RsProducto!lCombinacion, -1, 0) & "," _
                   & IIf(IsNull(RsProducto!nCombinacion), 0, RsProducto!nCombinacion) & "," _
                   & "'N', '" & sComanda & "', getDate(), '0000', '" & sUsuario & "', " _
                   & nInsumo & ", " & nGasto & ", " & nMObra & ", " _
                   & "1, '" & sUnidadNegocio & "','','','" & Format(obtieneDiaContable, "yyyyMMdd") & "', '" & sCaja & "','" & Trim(txtAutorizacion.Text) & "')"
        Cn.Execute Isql
    Else

    End If
    
'Cn.CommitTrans
InsertarPedidoAnticipo = sPedido
Exit Function

fin:
'Cn.RollbackTrans
InsertarPedidoAnticipo = ""
MsgBox "ERROR: " & error
End Function

'diaContable
Public Function obtieneDiaContable() As Date
   Dim oComando As New clsComando
   Dim DiaContable As Date
   Dim rst1 As New ADODB.Recordset
   Set oComando = New clsComando
                  If Not oComando.CreateCmdSp("usp_GenObtieneDiaContable", Cn) Then
                     Set oComando = Nothing
                     Exit Function
                  End If
                  
                  oComando.CreateParameter "@lDiaContable", adBoolean, adParamInput, 1, lDiaContable
                  oComando.CreateParameter "@sHoraCierre", adVarChar, adParamInput, 5, tHoraCierreDiaContable
                  oComando.CreateParameter "@tUsuario", adVarChar, adParamInput, 15, sUsuario
                 oComando.CreateParameter "@fDiaContable", adDBDate, adParamOutput, 10, DiaContable
                If Not oComando.GetParamOK Then
                   Set oComando = Nothing
                   Exit Function
                End If
                    Set rst1 = oComando.GetSP()
                obtieneDiaContable = oComando.GetParameterValue("@fDiaContable")
End Function

Private Sub txtCodigoTarjeta_KeyPress(KeyAscii As Integer)
'2A5545d2E55245
Dim Obtener As New ADODB.Recordset
        If KeyAscii = 13 Then  ' The ENTER key.
'            SendKeys "{tab}"    ' Set the focus to the next control.
'            KeyAscii = 0        ' Ignore this key.
            Isql = "select CodidoRFID codigo, CodigoCliente  , t2.tApellido + ' ' + t2.tNombre descripcion , Estado , isnull(t1.MontoDisponible,0) MontoDisponible     from TTARJETASRFID t1 inner join TDELIVERY t2 on t1.CodigoCliente = t2.tCodigoDelivery where CodidoRFID =  '" & txtCodigoTarjeta.Text & "'"
            Set Obtener = Lib.OpenRecordset(Isql, Cn)
            If Obtener.BOF Or Obtener.EOF Then
                LblCliente.Caption = "*** Cliente ***"
                LblSaldo.Caption = Format(0, "###,###,###,##0.00")
                MsgBox "Tarjeta no existe, verificar datos ingresados!!", vbCritical, sMensaje
                 
            Else
                 LblCliente.Caption = IIf(IsNull(Obtener!Descripcion), 0, Obtener!Descripcion)
                LblSaldo.Caption = Format(IIf(IsNull(Obtener!MontoDisponible), 0, Obtener!MontoDisponible), "###,###,###,##0.00")
            End If
         End If


End Sub
