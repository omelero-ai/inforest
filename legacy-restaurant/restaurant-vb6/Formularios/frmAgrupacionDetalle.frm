VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAgrupacionDetalle 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6690
   ClientLeft      =   -570
   ClientTop       =   1875
   ClientWidth     =   9495
   FillColor       =   &H8000000F&
   Icon            =   "frmAgrupacionDetalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9495
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
      Height          =   5910
      Left            =   1710
      TabIndex        =   31
      Top             =   45
      Width           =   7755
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
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1170
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
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   630
         Width           =   4935
      End
      Begin VB.Frame fraGrilla 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4830
         Left            =   135
         TabIndex        =   22
         Top             =   990
         Width           =   7515
         Begin VB.CommandButton cmdBoton 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   6390
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   270
            Width           =   960
         End
         Begin VB.CommandButton cmdBoton 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   9
            Left            =   6390
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   4230
            Width           =   960
         End
         Begin VB.CommandButton cmdBoton 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   8
            Left            =   6390
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   3735
            Width           =   960
         End
         Begin VB.CommandButton cmdBoton 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   7
            Left            =   6390
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   3240
            Width           =   960
         End
         Begin VB.CommandButton cmdBoton 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   6
            Left            =   6390
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   2745
            Width           =   960
         End
         Begin VB.CommandButton cmdBoton 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   5
            Left            =   6390
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   2250
            Width           =   960
         End
         Begin VB.CommandButton cmdBoton 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   4
            Left            =   6390
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1755
            Width           =   960
         End
         Begin VB.CommandButton cmdBoton 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   3
            Left            =   6390
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1260
            Width           =   960
         End
         Begin VB.CommandButton cmdBoton 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   2
            Left            =   6390
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   765
            Width           =   960
         End
         Begin VB.TextBox txtDetalle 
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
            Index           =   8
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   9
            Text            =   " "
            Top             =   3810
            Width           =   4890
         End
         Begin VB.TextBox txtDetalle 
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
            Index           =   7
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   8
            Text            =   " "
            Top             =   3315
            Width           =   4890
         End
         Begin VB.TextBox txtDetalle 
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
            Index           =   6
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   7
            Text            =   " "
            Top             =   2820
            Width           =   4890
         End
         Begin VB.TextBox txtDetalle 
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
            Index           =   5
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   6
            Text            =   " "
            Top             =   2325
            Width           =   4890
         End
         Begin VB.TextBox txtDetalle 
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
            Index           =   4
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   5
            Text            =   " "
            Top             =   1830
            Width           =   4890
         End
         Begin VB.TextBox txtDetalle 
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
            Index           =   3
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   4
            Text            =   " "
            Top             =   1335
            Width           =   4890
         End
         Begin VB.TextBox txtDetalle 
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
            Index           =   2
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   3
            Text            =   " "
            Top             =   840
            Width           =   4890
         End
         Begin VB.TextBox txtDetalle 
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
            Index           =   1
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   2
            Text            =   " "
            Top             =   345
            Width           =   4890
         End
         Begin VB.TextBox txtDetalle 
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
            Index           =   9
            Left            =   1350
            MaxLength       =   50
            TabIndex        =   10
            Text            =   " "
            Top             =   4275
            Width           =   4890
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Agrupación 9 :"
            Height          =   195
            Index           =   10
            Left            =   225
            TabIndex        =   43
            Top             =   4350
            Width           =   1035
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Agrupación 8 :"
            Height          =   195
            Index           =   9
            Left            =   225
            TabIndex        =   42
            Top             =   3855
            Width           =   1035
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Agrupación 7 :"
            Height          =   195
            Index           =   8
            Left            =   225
            TabIndex        =   41
            Top             =   3360
            Width           =   1035
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Agrupación 6 :"
            Height          =   195
            Index           =   7
            Left            =   225
            TabIndex        =   40
            Top             =   2865
            Width           =   1035
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Agrupación 5 :"
            Height          =   195
            Index           =   6
            Left            =   225
            TabIndex        =   39
            Top             =   2370
            Width           =   1035
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Agrupación 4 :"
            Height          =   195
            Index           =   5
            Left            =   225
            TabIndex        =   38
            Top             =   1875
            Width           =   1035
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Agrupación 3 :"
            Height          =   195
            Index           =   3
            Left            =   225
            TabIndex        =   37
            Top             =   1380
            Width           =   1035
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Agrupación 2 :"
            Height          =   195
            Index           =   2
            Left            =   225
            TabIndex        =   36
            Top             =   885
            Width           =   1035
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Agrupación 1 :"
            Height          =   195
            Index           =   4
            Left            =   225
            TabIndex        =   35
            Top             =   383
            Width           =   1035
         End
      End
      Begin MSComDlg.CommonDialog cmdColor 
         Left            =   7155
         Top             =   450
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Código :"
         Height          =   195
         Index           =   0
         Left            =   765
         TabIndex        =   33
         Top             =   285
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Etiqueta :"
         Height          =   195
         Index           =   1
         Left            =   675
         TabIndex        =   32
         Top             =   675
         Width           =   675
      End
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   9435
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5985
      Width           =   9495
      Begin VB.PictureBox PicNavegacion 
         BackColor       =   &H80000004&
         Height          =   645
         Left            =   30
         ScaleHeight     =   585
         ScaleWidth      =   6810
         TabIndex        =   24
         Top             =   60
         Width           =   6870
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   3
            Left            =   5370
            Picture         =   "frmAgrupacionDetalle.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   4
            Left            =   5850
            Picture         =   "frmAgrupacionDetalle.frx":0984
            Style           =   1  'Graphical
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   5
            Left            =   6330
            Picture         =   "frmAgrupacionDetalle.frx":0EC6
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   0
            Left            =   0
            Picture         =   "frmAgrupacionDetalle.frx":1408
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   2
            Left            =   960
            Picture         =   "frmAgrupacionDetalle.frx":194A
            Style           =   1  'Graphical
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   0
            Width           =   480
         End
         Begin VB.CommandButton cmdNavegar 
            Height          =   555
            Index           =   1
            Left            =   480
            Picture         =   "frmAgrupacionDetalle.frx":1E8C
            Style           =   1  'Graphical
            TabIndex        =   25
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
            Left            =   1485
            TabIndex        =   34
            Top             =   150
            Width           =   3630
         End
      End
      Begin VB.CommandButton cmdGrabar 
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
         Left            =   7035
         Picture         =   "frmAgrupacionDetalle.frx":23CE
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   15
         Width           =   1170
      End
      Begin VB.CommandButton cmdSalir 
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
         Left            =   8250
         Picture         =   "frmAgrupacionDetalle.frx":2900
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   15
         Width           =   1170
      End
   End
   Begin VB.Image Image 
      Height          =   5850
      Left            =   30
      Picture         =   "frmAgrupacionDetalle.frx":29F2
      Stretch         =   -1  'True
      Top             =   90
      Width           =   1620
   End
End
Attribute VB_Name = "frmAgrupacionDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sprefijo As String
Dim i As Integer
Dim sCodigo As String
Dim RsDetalle As ADODB.Recordset

Private Sub cmdBoton_Click(Index As Integer)
   On Error GoTo cmdCancel
   cmdColor.CancelError = True
   cmdColor.ShowColor
   cmdBoton(Index).BackColor = cmdColor.Color
Exit Sub
cmdCancel:
   cmdBoton(Index).BackColor = -2147483633
End Sub

Private Sub cmdGrabar_Click()
   nPos = frmAgrupacion.RsCabecera.Bookmark
   Cn.Execute "update TTABLA set tDetallado='" & txtDetallado.Text & "' where tTabla='ETIQUETA' and tCodigo='0" & sprefijo & "'"
   For i = 1 To 9
       sCodigo = sprefijo & LTrim(Str(i))
       Cn.Execute "update TTABLA set tDetallado='" & txtDetalle(i).Text & "', nValor= " & cmdBoton(i).BackColor & " where tTabla='CAJARAPIDA' and tCodigo='" & sCodigo & "'"
   Next i
   RsDetalle.Requery
   frmAgrupacion.RsCabecera.Requery
   frmAgrupacion.RsCabecera.Bookmark = nPos
   Screen.MousePointer = vbDefault
   MsgBox "Registro Modificado", vbInformation, sMensaje
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
    Centrar Me
    Me.Caption = "Mantenimiento de Agrupación de Caja Rápida"
    Set RsDetalle = Lib.OpenRecordset("select tCodigo, substring(tCodigo,1,1) as Prefijo, tDetallado, nValor from TTABLA where tTabla='CAJARAPIDA'", Cn)
    Asignar
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmAgrupacionDetalle = Nothing
End Sub

Private Sub txtDetalle_LostFocus(Index As Integer)
   cmdBoton(Index).Caption = txtDetalle(Index).Text
End Sub

Private Sub cmdNavegar_Click(Index As Integer)
    Select Case Index
           Case Is = 0 'Primero
                MoverPuntero Primero, frmAgrupacion.grdGrilla
           Case Is = 1 'PgUp
                MoverPuntero pgup, frmAgrupacion.grdGrilla
           Case Is = 2 'Previo
                MoverPuntero previo, frmAgrupacion.grdGrilla
           Case Is = 3 'Siguiente
                MoverPuntero siguiente, frmAgrupacion.grdGrilla
           Case Is = 4 'PgDn
                MoverPuntero pgdn, frmAgrupacion.grdGrilla
           Case Is = 5 'Ultimo
                MoverPuntero Ultimo, frmAgrupacion.grdGrilla
    End Select
    Asignar
End Sub

Sub Asignar()
    With frmAgrupacion.RsCabecera
        'Cuadro de Texto
        txtCodigo = IIf(IsNull(!Codigo), "", !Codigo)
        txtDetallado = IIf(IsNull(!Descripcion), "", !Descripcion)
        sprefijo = Mid(txtCodigo, 2, 1)
        RsDetalle.Filter = "Prefijo='" & sprefijo & "'"
        RsDetalle.MoveFirst
        For i = 1 To 9
            sCodigo = sprefijo & LTrim(Str(i))
            cmdBoton(i).BackColor = IIf(RsDetalle!nValor = 0, -2147483633, RsDetalle!nValor)
            txtDetalle(i).Text = RsDetalle!tDetallado
            cmdBoton(i).Caption = RsDetalle!tDetallado
            RsDetalle.MoveNext
        Next i

    End With
    cmdTexto.Caption = "Registro " & IIf(frmAgrupacion.RsCabecera.RecordCount = 0, 0, frmAgrupacion.RsCabecera.AbsolutePosition) & " de " & frmAgrupacion.RsCabecera.RecordCount
End Sub
