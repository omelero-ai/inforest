Attribute VB_Name = "modDeclaraciones"
Option Explicit

Public Valor As String

Public ResY As String
Public ResM As String
Public ResD As String
Public ResH As String
Public ResN As String

Public Cn As Connection
Public Rs As Recordset

Public unidad(0 To 9) As String
Public decena(0 To 9) As String
Public centena(0 To 10) As String
Public deci(0 To 9) As String
Public otros(0 To 15) As String
Public inga As Double
Public negativo As Boolean
Public l As Integer
Public un As Boolean, millon As Boolean, millones As Boolean
Public vez As Integer
Public maxvez As Integer
Public k As Integer
Public cola As String
Public strq As String, strb As String, stru As String
Public strd As String, strc As String
Public ia As Integer
Public strn() As String
Public cTabla As String
Public Campo1 As String
Public Campo2 As String
Public Conect As Connection
