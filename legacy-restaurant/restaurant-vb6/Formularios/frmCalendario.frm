VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCalendario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendario"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   FillColor       =   &H80000012&
   Icon            =   "frmCalendario.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Mes Previo"
      Height          =   585
      Index           =   3
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4830
      Width           =   1365
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "Mes Siguiente"
      Height          =   585
      Index           =   2
      Left            =   1455
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4830
      Width           =   1365
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
      Left            =   2850
      Picture         =   "frmCalendario.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4830
      Width           =   1365
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
      Left            =   4245
      Picture         =   "frmCalendario.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4830
      Width           =   1365
   End
   Begin MSComCtl2.MonthView Calendar 
      Height          =   4770
      Left            =   30
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   8414
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   56885249
      CurrentDate     =   37474
   End
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0  'Aceptar
               wEnter = True
               sDescrip = Calendar.Value
               Unload Me
          
          Case Is = 1  'Cancelar
               Unload Me
          
          Case Is = 2  'Mes Siguiente
               If Calendar.Month = 12 Then
                  Calendar.Month = 1
                  Calendar.Year = Calendar.Year + 1
               Else
                  Calendar.Month = Calendar.Month + 1
               End If
          
          Case Is = 3  'Mes Previo
               If Calendar.Month = 1 Then
                  Calendar.Month = 12
                  Calendar.Year = Calendar.Year - 1
               Else
                  Calendar.Month = Calendar.Month - 1
               End If
   End Select
End Sub

Private Sub Form_Load()
   Centrar Me
   wEnter = False
   Calendar.Value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmCalendario = Nothing
End Sub
