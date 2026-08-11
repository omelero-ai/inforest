VERSION 5.00
Begin VB.Form frmFactura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Correlativo de Documentos"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   Icon            =   "frmFactura.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCorrelativo 
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
      Height          =   315
      Index           =   2
      Left            =   2040
      TabIndex        =   8
      Top             =   1650
      Width           =   1230
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Correlativo"
      Height          =   585
      Index           =   3
      Left            =   3330
      TabIndex        =   7
      Top             =   1530
      Width           =   1230
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Correlativo"
      Height          =   585
      Index           =   2
      Left            =   3330
      TabIndex        =   5
      Top             =   675
      Width           =   1230
   End
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
      Height          =   585
      Index           =   1
      Left            =   2010
      Picture         =   "frmFactura.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2550
      Width           =   1230
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
      Height          =   585
      Index           =   0
      Left            =   3330
      Picture         =   "frmFactura.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2550
      Width           =   1230
   End
   Begin VB.TextBox txtCorrelativo 
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
      Height          =   315
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   810
      Width           =   1230
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Correlativo de Factura :"
      Height          =   195
      Index           =   4
      Left            =   90
      TabIndex        =   10
      Top             =   1740
      Width           =   1890
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "(Ultimo Generado)"
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
      Index           =   3
      Left            =   195
      TabIndex        =   9
      Top             =   2010
      Width           =   1785
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "(Ultimo Generado)"
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
      Index           =   2
      Left            =   195
      TabIndex        =   6
      Top             =   1170
      Width           =   1785
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "El mal cambio de Correlativo Puede traer Errores de Generación de Documento. ¡CUIDADO!"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   4425
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Correlativo de Factura :"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   0
      Top             =   870
      Width           =   1890
   End
End
Attribute VB_Name = "frmFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsFactura As Recordset
Dim vFlag(1 To 2) As Boolean


Private Sub cmdOpcion_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Aceptar
                If vFlag(1) Then
                   Cn.Execute ("update TTIPODOCUMENTOIMPRESORA set tUltimoNumero='" & txtCorrelativo(1).Text & "' where tCaja='" & sCaja & "' and tTipoEmision='01'")
                End If
                
                If vFlag(2) Then
                   Cn.Execute ("update TTIPODOCUMENTOIMPRESORA set tUltimoNumero='" & txtCorrelativo(2).Text & "' where tCaja='" & sCaja & "' and tTipoEmision='02'")
                End If
                
                NFactura = txtCorrelativo(1).Text
                Unload Me
                
           Case Is = 1 'Cancelar
                wEnter = False
                Unload Me
                
           Case Is = 2 'Numpad
                sTipo = "Numero"
                frmNumPad.Show vbModal
                If wEnter Then
                   txtCorrelativo(1).Text = Mid("000000000", 1, 9 - Len(Trim(sDescrip))) & Trim(sDescrip)
                End If
           
           Case Is = 3 'Numpad
                sTipo = "Numero"
                frmNumPad.Show vbModal
                If wEnter Then
                   txtCorrelativo(2).Text = Mid("000000000", 1, 9 - Len(Trim(sDescrip))) & Trim(sDescrip)
                End If
    
    End Select
End Sub

Private Sub Form_Load()
   Centrar Me
   Me.Caption = "Correlativo de Factura Caja : " & sCaja
   Isql = "select * from vTIPODOCUMENTOIMPRESORA where tCaja ='" & sCaja & "' and Transporte=0"
   Set RsFactura = Lib.OpenRecordset(Isql, Cn)
   If Not (RsFactura.EOF Or RsFactura.BOF) Then
        RsFactura.MoveFirst
        RsFactura.Find ("tTipoEmision = '01'")
             If Not (RsFactura.EOF) Then
                 Label(1).Caption = "Correlativo de " & RsFactura!Descripcion
                 txtCorrelativo(1).Text = RsFactura!tUltimoNumero
                 vFlag(1) = True
             End If
        RsFactura.MoveFirst
   Else
        cmdOpcion(2).Enabled = False
        vFlag(1) = False
   End If
   
   If Not (RsFactura.EOF Or RsFactura.BOF) Then
        RsFactura.MoveFirst
        RsFactura.Find ("tTipoEmision = '02'")
             If Not (RsFactura.EOF) Then
                 Label(4).Caption = "Correlativo de " & RsFactura!Descripcion
                 txtCorrelativo(2).Text = RsFactura!tUltimoNumero
                 vFlag(2) = True
             End If
        RsFactura.MoveFirst
   Else
        cmdOpcion(3).Enabled = False
        vFlag(2) = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFactura = Nothing
End Sub
