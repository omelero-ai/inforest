VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Encripta una cadena de caracteres.
'S = Cadena a encriptar
'P = Password
Public Function TextEncript(ByVal S As String, Optional ByVal P As String = "123456789abcde") As String
Dim i As Integer, r As String
Dim C1 As Integer, C2 As Integer
r = ""
If Len(P) > 0 Then
For i = 1 To Len(S)
C1 = Asc(Mid(S, i, 1))
If i > Len(P) Then
C2 = Asc(Mid(P, i Mod Len(P) + 1, 1))
Else
C2 = Asc(Mid(P, i, 1))
End If
C1 = C1 + C2 + 64
If C1 > 255 Then C1 = C1 - 256
r = r + Chr(C1)
Next i
Else
r = S
End If
TextEncript = r

End Function

'Desencripta una cadena de caracteres.
'S = Cadena a desencriptar
'P = Password
Public Function TextDecript(ByVal S As String, Optional ByVal P As String = "123456789abcde") As String
Dim i As Integer, r As String
Dim C1 As Integer, C2 As Integer
r = ""
If Len(P) > 0 Then
For i = 1 To Len(S)
C1 = Asc(Mid(S, i, 1))
If i > Len(P) Then
C2 = Asc(Mid(P, i Mod Len(P) + 1, 1))
Else
C2 = Asc(Mid(P, i, 1))
End If
C1 = C1 - C2 - 64
If Sgn(C1) = -1 Then C1 = 256 + C1
r = r + Chr(C1)
Next i
Else
r = S
End If
TextDecript = r
End Function
