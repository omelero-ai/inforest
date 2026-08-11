VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFecha 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rango de Fechas"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   Icon            =   "frmFecha.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   420
      Picture         =   "frmFecha.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1020
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   1695
      Picture         =   "frmFecha.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1020
      Width           =   1275
   End
   Begin MSComCtl2.DTPicker dtpFecha2 
      Height          =   345
      Left            =   1170
      TabIndex        =   2
      Top             =   465
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   609
      _Version        =   393216
      Format          =   24576001
      CurrentDate     =   38180
   End
   Begin MSComCtl2.DTPicker dtpFecha1 
      Height          =   345
      Left            =   1170
      TabIndex        =   1
      Top             =   60
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   24576001
      CurrentDate     =   38180
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Final :"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   3
      Top             =   540
      Width           =   915
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Inicial :"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   135
      Width           =   990
   End
End
Attribute VB_Name = "frmFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOpcion_Click(Index As Integer)
   If dtpFecha1.Value >= dtpFecha2.Value Then
      MsgBox "Error en rango de fechas", vbCritical, sMensaje
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   If Index = 0 Then
      fFechaI = dtpFecha1.Value
      fFechaF = dtpFecha2.Value
      lCopia = False
      nMonto = Calcular("select sum(nVenta) as codigo FROM dbo.vDocumentoGrilla LEFT OUTER JOIN " & _
                        "(select DDOCUMENTO.tDocumento, count(distinct tCodigoPedido) as TotalDocumento from dbo.DDOCUMENTO LEFT OUTER JOIN dbo.MDOCUMENTO ON dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento where fRegistro >= '" & Format(fFechaI, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(fFechaF, "yyyy/MM/dd") & "' and tTipodocumento='02' group by DDOCUMENTO.tDocumento) T1 " & _
                        "ON dbo.vDocumentoGrilla.tDocumento = T1.tDocumento LEFT OUTER JOIN " & _
                        "(select tCodigoPedido, count(distinct DDOCUMENTO.tDocumento) as TotalPedido from dbo.DDOCUMENTO LEFT OUTER JOIN dbo.MDOCUMENTO ON dbo.DDOCUMENTO.tDocumento = dbo.MDOCUMENTO.tDocumento where fRegistro >= '" & Format(fFechaI, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(fFechaF, "yyyy/MM/dd") & "' and tTipodocumento='02' group by tCodigoPedido) T2 " & _
                        "ON dbo.vDocumentoGrilla.tCodigoPedido = T2.tCodigoPedido " & _
                        "where fFecha >= '" & Format(fFechaI, "yyyy/MM/dd") & "' and fFecha <= '" & Format(fFechaF, "yyyy/MM/dd") & "' and tTipodocumento='02' and tEstadoDocumento='02' and T1.TotalDocumento <= 1 and T2.TotalPedido <= 1 and vDocumentoGrilla.tDocumento not in " & _
                        "(select tDocumento from DPAGODOCUMENTO where (tTipoPago='02' OR tTipoPago='03' OR tTipoPago='04' OR tTipoPago='05') and fRegistro >= '" & Format(fFechaI, "yyyy/MM/dd") & "' and fRegistro <= '" & Format(fFechaF, "yyyy/MM/dd") & "')", Cn)
                         
      Cn.Execute "Insert MENVIO (fInicio, fFinal, tUsuario, nMonto, lCopia, lCierre) values ('" & Format(fFechaI, "yyyy/MM/dd") & "', '" & Format(fFechaF, "yyyy/MM/dd") & "', 'MASTER', " & nMonto & ", 0, 0)"
      Screen.MousePointer = vbDefault
      Unload Me
   Else
      End
   End If

End Sub

Private Sub Form_Load()
   dtpFecha2.Value = Date
End Sub
