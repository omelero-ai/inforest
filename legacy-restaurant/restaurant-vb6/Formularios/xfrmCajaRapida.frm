VERSION 5.00
Begin VB.Form frmCajaRapida 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCajaRapida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sDetalle As String
Public sPropiedad As String
Public sComboDetalle As String
Public sComboPropiedad As String
Public sProductoPropiedad As String
Public Pedido As String

