VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMsComm 
   Caption         =   "Formulario de Comunicaciones"
   ClientHeight    =   765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   ScaleHeight     =   765
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin MSCommLib.MSComm MSCommPort 
      Left            =   135
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      ParitySetting   =   2
      DataBits        =   7
   End
   Begin MSCommLib.MSComm MSCommBalanza 
      Left            =   840
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      ParitySetting   =   2
      DataBits        =   7
   End
End
Attribute VB_Name = "frmMsComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MSCommBalanza_OnComm()
'        Dim InBuffer As String
'            InBuffer = MSCommBalanza.Input
'         If MSCommBalanza.CommEvent = comEvReceive Then
'            'Text2.Text = Text2.Text & InBuffer
'         End If
End Sub

