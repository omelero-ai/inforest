VERSION 5.00
Begin VB.Form frmJuntaMesa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Juntar Mesas"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   7275
   FillColor       =   &H8000000B&
   Icon            =   "frmJuntaMesa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   7275
   Begin VB.Frame fraMesa 
      Caption         =   " Mesas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8235
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   5715
      Begin VB.CommandButton cmdMesa 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   540
         Width           =   915
      End
      Begin VB.Image imgFondo 
         Height          =   7815
         Left            =   120
         Stretch         =   -1  'True
         Top             =   330
         Width           =   5505
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mesas Juntas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   5760
      TabIndex        =   12
      Top             =   5130
      Width           =   1455
      Begin VB.TextBox txtJuntas 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   225
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   315
         Width           =   1005
      End
   End
   Begin VB.Frame fraOpcion 
      Caption         =   " Opciones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   5760
      TabIndex        =   7
      Top             =   5985
      Width           =   1455
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
         Height          =   555
         Index           =   3
         Left            =   90
         Picture         =   "frmJuntaMesa.frx":2582
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1545
         Width           =   1275
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
         Height          =   555
         Index           =   1
         Left            =   90
         Picture         =   "frmJuntaMesa.frx":2674
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   900
         Width           =   1275
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Refrescar"
         Height          =   555
         Index           =   2
         Left            =   90
         TabIndex        =   9
         Top             =   255
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
         Left            =   90
         Picture         =   "frmJuntaMesa.frx":2776
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1545
         Width           =   1275
      End
   End
   Begin VB.Frame fraEstado 
      Caption         =   " Leyenda "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3870
      Left            =   5760
      TabIndex        =   0
      Top             =   45
      Width           =   1455
      Begin VB.CommandButton cmdEstado 
         BackColor       =   &H00FF8080&
         Caption         =   "Lista"
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
         Height          =   555
         Index           =   1
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1275
      End
      Begin VB.CommandButton cmdEstado 
         BackColor       =   &H0080C0FF&
         Caption         =   "Ocupada"
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
         Height          =   555
         Index           =   2
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   831
         Width           =   1275
      End
      Begin VB.CommandButton cmdEstado 
         BackColor       =   &H00008000&
         Caption         =   "Reservada"
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
         Height          =   555
         Index           =   3
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2013
         Width           =   1275
      End
      Begin VB.CommandButton cmdEstado 
         BackColor       =   &H0000C0C0&
         Caption         =   "Sucia"
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
         Height          =   555
         Index           =   4
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1422
         Width           =   1275
      End
      Begin VB.CommandButton cmdEstado 
         BackColor       =   &H000000C0&
         Caption         =   "Bloqueada"
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
         Height          =   555
         Index           =   5
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2604
         Width           =   1275
      End
      Begin VB.CommandButton cmdEstado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fuera de Servicio"
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
         Height          =   555
         Index           =   6
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3195
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmJuntaMesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsMesa As Recordset
Dim RsSalon As Recordset
Dim RsPedidoMesa As Recordset
Dim sJuntaMesa() As String
Dim nCantidad As Integer
Public sSalonActual As String

Private Sub cmdMesa_Click(Index As Integer)
     If sModulo = "ADICION" Then
        frmMozoUsuario.ReseteaTimer
     End If

     RsMesa.MoveFirst
     RsMesa.Move (Index - 1)
     
     If cmdMesa(Index).BackColor = vbFServicio Then
        cmdMesa(Index).BackColor = vbLista
        nCantidad = nCantidad - 1
        sJuntaMesa(Index, 3) = "00"
     Else
        cmdMesa(Index).BackColor = vbFServicio
        nCantidad = nCantidad + 1
        sJuntaMesa(Index, 3) = "06"
     End If
     txtJuntas.Text = Format(nCantidad, "##0")
               
End Sub

Private Sub cmdOpcion_Click(Index As Integer)
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
Dim i As Integer
  Select Case Index
         Case Is = 0 ' Aceptar
'              If nCantidad < 1 Then
'                 MsgBox "Debes Juntar al menos una mesa", vbExclamation, sMensaje
'                 Exit Sub
'              End If
              
              'Consistenciar Primero
              Dim sEstado As String
              For i = 1 To RsMesa.RecordCount
                  If sJuntaMesa(i, 3) = "06" Then
                     sEstado = Calcular("select tEstadoMesa as Codigo from TMESA where tCodigoMesa='" & sJuntaMesa(i, 1) & "'", Cn)
                     If (sEstado = "02" Or sEstado = "03" Or sEstado = "05" Or sEstado = "06") And sJuntaMesa(i, 3) <> "06" Then
                        MsgBox "Error: Has Juntado una Mesa Ocupada, verifique", vbExclamation, sMensaje
                        Exit Sub
                     End If
                  End If
              Next i
              
              Cn.Execute "update TMESA set tEstadoMesa='01' where tCodigoMesa in (select tMesa from TPEDIDOMESA where tCodigoPedido='" & sPedido & "')"
              Cn.Execute "delete from TPEDIDOMESA where tCodigoPedido='" & sPedido & "'"
              
              
              RsMesa.MoveFirst
              For i = 1 To RsMesa.RecordCount
                 If sJuntaMesa(i, 3) = "06" Then
                    Cn.Execute "INSERT into TPEDIDOMESA select '" & sPedido & "','" & sJuntaMesa(i, 1) & "'"
                    Cn.Execute "update TMESA set tEstadoMesa='06' where tCodigoMesa='" & sJuntaMesa(i, 1) & "'"
                 End If
              Next i
              Unload Me
              nVar1 = nCantidad
              wEnter = True
              
         Case Is = 1 ' Cancelar
              wEnter = False
              Unload Me

         Case Is = 2 ' Refrescar
              RsMesa.Requery
              Desasigna (RsMesa.RecordCount)
              AsignaMesa
              RsMesa.MoveFirst
              Set RsPedidoMesa = Lib.OpenRecordset("select tMesa From TPEDIDOMESA where tCodigoPedido='" & sPedido & "'", Cn)
              nCantidad = RsPedidoMesa.RecordCount
              txtJuntas.Text = Format(nCantidad, "##0")
              
         Case Is = 3 ' Salir
              Unload Me

  End Select
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Centrar Me
   If sModulo = "ADICION" Then
      frmMozoUsuario.ReseteaTimer
   End If
   
   Set RsMesa = Lib.OpenRecordset("select * from TMESA where tCodigoMesa <> '000' and lActivo= 1 and tSalon='" & sSalonActual & "' order by tCodigoMesa", Cn)
   Set RsSalon = Lib.OpenRecordset("select tCodigo as tSalon, tDetallado as dSalon, tiCono from TTABLA where tTABLA ='SALON' and lActivo = 1 and tCodigo='" & sSalonActual & "'", Cn)
   Set RsPedidoMesa = Lib.OpenRecordset("select tMesa From TPEDIDOMESA where tCodigoPedido='" & sPedido & "'", Cn)
   nCantidad = RsPedidoMesa.RecordCount
   txtJuntas.Text = Format(nCantidad, "##0")
      
   ActivaOpcion True
   
   fraMesa.Caption = " " & RsSalon!dSalon & " "
   InsertaBMP IIf(IsNull(RsSalon!tIcono), "", RsSalon!tIcono)
   AsignaMesa
   Screen.MousePointer = vbDefault
End Sub

Public Sub AsignaMesa()
   Dim i As Integer
   Dim nHora As Long
   
   RsMesa.MoveFirst
   cmdMesa(0).Visible = False
   ReDim sJuntaMesa(RsMesa.RecordCount, 3)
   
   For i = 1 To RsMesa.RecordCount
       'Cantidad de Mesas
       With cmdMesa(i)
            Load cmdMesa(i)
            .Top = RsMesa!tX
            .Left = RsMesa!tY
            .Visible = True
            cmdMesa(i).Caption = Trim(RsMesa!tResumido)
            sJuntaMesa(i, 1) = Trim(RsMesa!tCodigoMesa)
            sJuntaMesa(i, 2) = Trim(RsMesa!tResumido)
            sJuntaMesa(i, 3) = "00"
            cmdMesa(i).Enabled = False
            
            Select Case RsMesa!tEstadoMesa
                   Case Is = "01" ' Lista
                        cmdMesa(i).Enabled = True
                        cmdMesa(i).BackColor = vbLista
                        
                   Case Is = "02" ' Ocupada
                        cmdMesa(i).BackColor = vbOcupada
 
                   Case Is = "03" ' Reservada
                        cmdMesa(i).BackColor = vbReservada
                        
                   Case Is = "04" ' Sucia
                        cmdMesa(i).BackColor = vbSucia
                        cmdMesa(i).Enabled = True
                        
                   Case Is = "05" ' Bloqueada
                        cmdMesa(i).BackColor = vbBloqueada
                                                
                   Case Is = "06" ' Fuera de Servicio
                        cmdMesa(i).BackColor = vbFServicio
                        If RsPedidoMesa.RecordCount > 0 Then
                           RsPedidoMesa.MoveFirst
                           RsPedidoMesa.Find ("tMesa='" & sJuntaMesa(i, 1) & "'")
                        
                           If sTipo = "M" And Not RsPedidoMesa.EOF Then
                              cmdMesa(i).Enabled = True
                              sJuntaMesa(i, 3) = "06"
                           End If
                        End If
                        
            End Select
            
            'Cambio de Mesa - Marron
            If sCodigo = sJuntaMesa(i, 1) Then
               cmdMesa(i).BackColor = &H4080&
               cmdMesa(i).Enabled = False
            End If
            
       End With
       RsMesa.MoveNext
   Next i
   
End Sub

Public Sub Desasigna(Cantidad As Integer)
   Dim i As Integer
   For i = 1 To Cantidad
       Unload cmdMesa(i)
   Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set RsMesa = Nothing
   Set RsSalon = Nothing
   Set frmJuntaMesa = Nothing
End Sub

Public Sub ActivaOpcion(Activa As Boolean)
    cmdOpcion(0).Visible = Activa
    cmdOpcion(1).Visible = Activa
    cmdOpcion(3).Visible = Not Activa
End Sub
