VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPapeCuadratura 
   Caption         =   "Cuadratura"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4905
   LinkTopic       =   "Form3"
   ScaleHeight     =   1830
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   1350
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Exportar"
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
         Index           =   2
         Left            =   1440
         Picture         =   "frmPapeCuadratura.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Exportar"
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
         Index           =   1
         Left            =   1440
         Picture         =   "frmPapeCuadratura.frx":0822
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   525
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   84279297
         CurrentDate     =   37539
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Exportar archivo del dia:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   2100
      End
   End
   Begin MSComDlg.CommonDialog cmdSave 
      Left            =   120
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image 
      Height          =   1875
      Left            =   0
      Picture         =   "frmPapeCuadratura.frx":1044
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1305
   End
End
Attribute VB_Name = "frmPapeCuadratura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TipoProceso As Integer
Private Sub cmdOpcion_Click(Index As Integer)
    Dim vTrama As String
    Set clsTramaFE = New clsTrama
    Dim Rs As Recordset
    Dim NameArchivo As String
    Select Case Index
        Case Is = 1
            '----------------------------------------------------Cuadratura------------------------------------------------------------------------
            Set oComando = New clsComando
            If Not oComando.CreateCmdSp("USP_FactVolcadoTCP", Cn) Then
               Set oComando = Nothing
               Exit Sub
            End If
            
            oComando.CreateParameter "@fecha", adDBDate, adParamInput, 10, dtpFecIni
            oComando.CreateParameter "@tipooper", adInteger, adParamInput, 10, 1
            
            If Not oComando.GetParamOK Then
               Set oComando = Nothing
               Exit Sub
            End If
            
            Set Rs = oComando.GetSP()
            If Rs.RecordCount = 0 Then
                MsgBox "No hay datos a Exportar!!!"
                Exit Sub
            End If

            NameArchivo = "PPL_" & sRUC & "_" & Format(dtpFecIni, "yyyyMMdd") & "." & Hour(DateTime.Now)
            cmdSave.Filter = "Archivos de texto Txt|*.txt"
            cmdSave.FileName = NameArchivo
            cmdSave.ShowSave
        
            If cmdSave.FileName = "" Then
               Exit Sub
            ElseIf Dir(cmdSave.FileName, vbDirectory) <> "" And NameArchivo = Replace(cmdSave.FileTitle, ".txt", "") Then
               If MsgBox("El Archivo ya esxite desea reemplazarlo?... ", vbYesNo) = vbNo Then
                    Exit Sub
               End If
               Kill (cmdSave.FileName)
            ElseIf NameArchivo <> Replace(cmdSave.FileTitle, ".txt", "") Then
                If MsgBox("El nombre del achivo no es correcto Segun Documentacion Paperlees!!!" & vbNewLine & "¿Desea Continuar?", vbOKCancel) = vbCancel Then
                    Exit Sub
                End If
                NameArchivo = Replace(cmdSave.FileTitle, ".txt", "")
                If Dir(cmdSave.FileName, vbDirectory) <> "" Then
                    If MsgBox("El Archivo ya esxite desea reemplazarlo?... ", vbYesNo) = vbNo Then
                        Exit Sub
                    End If
                    Kill (cmdSave.FileName)
                End If
        
            End If
            vTrama = ConvertStringToUtf8String(clsTramaFE.VolcarTrama(Rs, 1))
            If CrearTxt(Replace(cmdSave.FileName, cmdSave.FileTitle, ""), NameArchivo, vTrama) Then
                MsgBox "Archivo Creado Satisfactoriamente!!!"
            Else
                MsgBox "Error en el proceso de generacion de archivo!!!"
            End If
            '-------------------------------------------------------------------------------------------------------------------------------
        Case Is = 2
            '----------------------------------------------------Baja------------------------------------------------------------------------
            Set oComando = New clsComando
            If Not oComando.CreateCmdSp("USP_FactVolcadoTCP", Cn) Then
               Set oComando = Nothing
               Exit Sub
            End If
            
            oComando.CreateParameter "@fecha", adDBDate, adParamInput, 10, dtpFecIni
            oComando.CreateParameter "@tipooper", adInteger, adParamInput, 10, 2
            
            If Not oComando.GetParamOK Then
               Set oComando = Nothing
               Exit Sub
            End If
            
            Set Rs = oComando.GetSP()
            If Rs.RecordCount = 0 Then
                MsgBox "No hay datos a Exportar!!!"
                Exit Sub
            End If
        
            
            'Dim NameArchivo As String
            NameArchivo = "Baja PPL_" & sRUC & "_" & Format(dtpFecIni, "yyyyMMdd") & "." & Hour(DateTime.Now)
            cmdSave.Filter = "Archivos de texto Txt|*.txt"
            cmdSave.FileName = NameArchivo
            cmdSave.ShowSave
        
            If cmdSave.FileName = "" Then
               Exit Sub
            ElseIf Dir(cmdSave.FileName, vbDirectory) <> "" And NameArchivo = Replace(cmdSave.FileTitle, ".txt", "") Then
               If MsgBox("El Archivo ya esxite desea reemplazarlo?... ", vbYesNo) = vbNo Then
                    Exit Sub
               End If
               Kill (cmdSave.FileName)
            ElseIf NameArchivo <> Replace(cmdSave.FileTitle, ".txt", "") Then
                If MsgBox("El nombre del achivo no es correcto Segun Documentacion Paperlees!!!" & vbNewLine & "¿Desea Continuar?", vbOKCancel) = vbCancel Then
                    Exit Sub
                End If
                NameArchivo = Replace(cmdSave.FileTitle, ".txt", "")
                If Dir(cmdSave.FileName, vbDirectory) <> "" Then
                    If MsgBox("El Archivo ya esxite desea reemplazarlo?... ", vbYesNo) = vbNo Then
                        Exit Sub
                    End If
                    Kill (cmdSave.FileName)
                End If
        
            End If
            vTrama = ConvertStringToUtf8String(clsTramaFE.VolcarTrama(Rs, 2))
            If CrearTxt(Replace(cmdSave.FileName, cmdSave.FileTitle, ""), NameArchivo, vTrama) Then
                MsgBox "Archivo Creado Satisfactoriamente!!!"
            Else
                MsgBox "Error en el proceso de generacion de archivo!!!"
            End If
            '--------------------------------------------------------------------------------------------------------------------------
     End Select
End Sub
Private Sub Form_Load()
    dtpFecIni.value = DateTime.Now
    If TipoProceso = 1 Then
        cmdOpcion(1).Visible = True
        cmdOpcion(2).Visible = False
        Me.Caption = "Cuadratura"
    Else
        cmdOpcion(1).Visible = False
        cmdOpcion(2).Visible = True
        Me.Caption = "Baja"
    End If
    'Label(0).Caption = "Exportar archivo del dia :" & Format(dtpFecIni, "dd/MM/yyyy")
End Sub
