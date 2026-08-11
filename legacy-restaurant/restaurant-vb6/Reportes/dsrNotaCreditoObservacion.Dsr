VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} dsrNotaCreditoObservacion 
   ClientHeight    =   14565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19665
   OleObjectBlob   =   "dsrNotaCreditoObservacion.dsx":0000
End
Attribute VB_Name = "dsrNotaCreditoObservacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Sub Section2_Format(ByVal pFormattingInfo As Object)
        
    Dim fso As Object
    Dim Cliente As StdPicture
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(App.path & "\cliente.jpg") Then
          Set ImagenCliente.FormattedPicture = LoadPicture(App.path & "\cliente.jpg")
    End If
    
End Sub


Private Sub Section4_Format(ByVal pFormattingInfo As Object)

    Dim fso1 As Object
    Dim Archivo As StdPicture
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FileExists(App.path & "\fact.bmp") Then
        Set ImagenHash.FormattedPicture = LoadPicture(App.path & "\fact.bmp")
    End If

End Sub
