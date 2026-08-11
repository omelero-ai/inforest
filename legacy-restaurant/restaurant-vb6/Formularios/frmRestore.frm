VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6C5FD781-9ED8-11D1-87C0-444553540000}#1.0#0"; "dzstactx.dll"
Object = "{0FB90DC1-97D1-11D1-87C0-444553540000}#1.0#0"; "duzactx.dll"
Begin VB.Form frmRestore 
   Caption         =   "Restore"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   Icon            =   "frmRestore.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   " Nota "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   45
      TabIndex        =   11
      Top             =   1530
      Width           =   5880
      Begin VB.Label Label1 
         Caption         =   "Se necesitará sacar a todos los usuarios del Sistema para realizar éste proceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   180
         TabIndex        =   12
         Top             =   180
         Width           =   5115
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Estado "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   30
      TabIndex        =   10
      Top             =   30
      Width           =   1035
      Begin VB.Image Image1 
         Height          =   615
         Left            =   210
         Picture         =   "frmRestore.frx":014A
         Stretch         =   -1  'True
         Top             =   300
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   615
         Left            =   210
         Picture         =   "frmRestore.frx":058C
         Stretch         =   -1  'True
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
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
      Height          =   1485
      Left            =   1080
      TabIndex        =   4
      Top             =   30
      Width           =   7950
      Begin VB.DriveListBox Drive 
         Height          =   315
         Left            =   2070
         TabIndex        =   14
         Top             =   1035
         Width           =   5250
      End
      Begin VB.TextBox txtFile 
         BackColor       =   &H80000014&
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
         Left            =   2070
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   218
         Width           =   5220
      End
      Begin VB.CommandButton cmdOpcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   7380
         Picture         =   "frmRestore.frx":09CE
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   203
         Width           =   480
      End
      Begin VB.TextBox txtPassword 
         BackColor       =   &H80000014&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2070
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   630
         Width           =   1815
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Unidad de Red Asignada :"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   13
         Top             =   1080
         Width           =   1875
      End
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Archivo :"
         Height          =   195
         Index           =   1
         Left            =   780
         TabIndex        =   9
         Top             =   255
         Width           =   630
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Password :"
         Height          =   195
         Index           =   0
         Left            =   630
         TabIndex        =   8
         Top             =   690
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Salir"
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
      Left            =   7605
      Picture         =   "frmRestore.frx":0B18
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1575
      Width           =   1440
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Ejecutar"
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
      Left            =   6075
      Picture         =   "frmRestore.frx":0C0A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1575
      Width           =   1440
   End
   Begin VB.Frame Frame1 
      Caption         =   " Barra de Estado "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   45
      TabIndex        =   0
      Top             =   2205
      Width           =   9000
      Begin DZSTACTXLibCtl.dzstactxctrl Barra 
         Height          =   240
         Left            =   75
         OleObjectBlob   =   "frmRestore.frx":0D54
         TabIndex        =   1
         Top             =   225
         Width           =   8850
      End
   End
   Begin MSComDlg.CommonDialog cdBrowse 
      Left            =   585
      Top             =   1170
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Abrir Archivo Existente"
      Filter          =   "Archivos ZIP |*.zip|Todos los Archivos |*.*"
   End
   Begin DUZACTXLibCtl.duzactxctrl unzip 
      Left            =   45
      OleObjectBlob   =   "frmRestore.frx":0DCC
      Top             =   1170
   End
End
Attribute VB_Name = "frmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sFile As String
Dim sRutaServidor As String

Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0
               Unload Me
          Case Is = 1
               If UCase(txtPassword) <> sPAdmin Then
                     MsgBox "Password Incorrecto", vbCritical, sMensaje
                     Exit Sub
               End If
            
               If txtFile.Text <> "" Then
                  Screen.MousePointer = vbHourglass
                  On Error GoTo ErrorRestore
                  Image2.Visible = True
                  Image1.Visible = False
                  Frame2.Refresh
                  sFile = Left(Right(txtFile.Text, 15), 11)
                  If Len(Drive.Drive) > 5 Then
                     sRutaServidor = Mid(Drive.Drive, 5, Len(Drive.Drive) - 5) & "\"
                  Else
                    sRutaServidor = Drive.Drive & "\"
                  End If
                  
                  unzip.NoDirectoryNamesFlag = True
                  unzip.MajorStatusFlag = True
                  unzip.ZIPFile = txtFile.Text
                  unzip.RecurseFlag = True
                  unzip.UpdateFlag = False
                  unzip.FreshenFlag = False

                  unzip.Destination = sRutaServidor
                  unzip.Filespec = "*.*"
                  unzip.ActionDZ = 8
                  
                  Cn.Execute "use MASTER"
                  If sTipo = "Local" Then
                     Cn.Execute "RESTORE DATABASE " & sMDB & " FROM DISK='" & sRutaServidor & sFile & "'"
                  Else
                     Cn.Execute "RESTORE DATABASE TEMPORAL FROM DISK='" & Mid(txtFile.Text, 1, Len(Trim(txtFile.Text)) - 4) & "'"
                  End If
                  
                  DeleteFile sRutaServidor & sFile
                  Image1.Visible = True
                  Image2.Visible = False
                  Screen.MousePointer = vbDefault
                  MsgBox "La Restauración se ha realizado con Exito" & Chr(13) _
                       & "          El sistema se reinicializará", vbInformation, "Restauración"
                  End
               Else
                  MsgBox "No existe el Archivo", vbCritical, "Restaurar Copias"
               End If
          
          Case Is = 2
               cdBrowse.DialogTitle = "Archivo Backup(Zip)"
               cdBrowse.FileName = txtFile.Text
               cdBrowse.ShowOpen
               txtFile.Text = cdBrowse.FileName
               
               If txtFile.Text <> "" Then
                  cmdOpcion(1).Enabled = True
                  txtFile.Enabled = True
                  txtFile.BackColor = &H80000014
                  txtPassword.Enabled = True
                  txtPassword.BackColor = &H80000014
               End If
   End Select
Exit Sub

ErrorRestore:
   Screen.MousePointer = vbDefault
   MsgBox err.Description & Chr(13) & "No se concreto con el restore", vbCritical, sMensaje
   DeleteFile sRutaServidor & sFile
End Sub

Private Sub Form_Load()
   Image2.Visible = False
   Centrar Me
   cdBrowse.InitDir = "C:\"
   
   cmdOpcion(1).Enabled = False
   txtFile.Enabled = False
   txtFile.BackColor = &H80000013
   txtPassword.Enabled = False
   txtPassword.BackColor = &H80000013
   Screen.MousePointer = vbDefault
   If sTipo = "Local" Then
      frmRestore.Caption = "Restore"
   Else
      frmRestore.Caption = "Restore Temporal"
   End If
End Sub

Private Sub unzip_UnZipMajorStatus(ItemName As String, Percent As Long, Cancel As Long)
   If (Barra.StatusText <> ItemName) Then
      Barra.StatusText = ItemName
   End If
   
   If (Barra.StatusPercent <> Percent) Then
      Barra.StatusPercent = Percent
   End If
End Sub
