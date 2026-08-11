VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBackup 
   Caption         =   "Backup"
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   Icon            =   "frmBackup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFile 
      BackColor       =   &H80000014&
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
      Height          =   315
      Left            =   2040
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   200
      Width           =   5220
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear BK"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      Width           =   1440
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   1215
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12489
            MinWidth        =   5292
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame3 
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
      Left            =   8640
      TabIndex        =   5
      Top             =   0
      Width           =   7950
      Begin VB.DriveListBox Drive 
         Height          =   315
         Left            =   2070
         TabIndex        =   1
         Top             =   1035
         Visible         =   0   'False
         Width           =   5250
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
         Picture         =   "frmBackup.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   630
         Visible         =   0   'False
         Width           =   480
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   345
         Left            =   2070
         TabIndex        =   0
         Top             =   225
         Visible         =   0   'False
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56295425
         CurrentDate     =   38184
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         Caption         =   "Unidad de Red Asignada :"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1095
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Backup :"
         Height          =   195
         Index           =   1
         Left            =   855
         TabIndex        =   6
         Top             =   300
         Visible         =   0   'False
         Width           =   1140
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
      Left            =   4320
      Picture         =   "frmBackup.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
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
      Left            =   6105
      Picture         =   "frmBackup.frx":0386
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1575
      Visible         =   0   'False
      Width           =   1440
   End
   Begin MSComDlg.CommonDialog cdBrowse 
      Left            =   5490
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Abrir Archivo Existente"
      Filter          =   "Archivos ZIP |*.zip|Todos los Archivos |*.*"
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1035
      Begin VB.Image imgVerde 
         Height          =   615
         Left            =   210
         Picture         =   "frmBackup.frx":04D0
         Stretch         =   -1  'True
         Top             =   300
         Width           =   615
      End
      Begin VB.Image imgRojo 
         Height          =   615
         Left            =   210
         Picture         =   "frmBackup.frx":0912
         Stretch         =   -1  'True
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Label label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Archivo :"
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   12
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sFile As String
Dim sRutaServidor As String
Dim sRuta As String
Dim RsBK As Recordset


Private Sub cmdOpcion_Click(Index As Integer)
   Select Case Index
          Case Is = 0  ' Salir
               Unload Me
               
          Case Is = 1  ' Ejecutar
               Cn.CommandTimeout = 0
               On Error GoTo err
               StatusBar.Panels(1).Text = "Iniciando.... "
               imgRojo.Visible = True
               imgVerde.Visible = False
               Frame2.Refresh
               StatusBar.Panels(1).Text = "Depurando....  "
               EliminaTemporal
               StatusBar.Panels(1).Text = "Compactando...."
               Screen.MousePointer = vbHourglass
               sFile = Left(Right(txtFile.Text, 15), 11)
               If Len(Drive.Drive) > 5 Then
                  sRutaServidor = Mid(Drive.Drive, 5, Len(Drive.Drive) - 5) & "\"
               Else
                 sRutaServidor = Drive.Drive & "\"
               End If
               Cn.Execute "use " & sMDB
               Cn.Execute "BACKUP DATABASE " & sMDB & " TO DISK='" & sRutaServidor & sFile & "'"
               StatusBar.Panels(1).Text = "Empaquetando..."
               
               Dim retcode As Integer  ' For Return Code From ZIP32.DLL
               zDate = vbNullString
               zJunkDir = 0     ' 1 = Throw Away Path Names
               zRecurse = 0     ' 1 = Recurse -r 2 = Recurse -R 2 = Most Useful :)
               zUpdate = 0      ' 1 = Update Only If Newer
               zFreshen = 0     ' 1 = Freshen - Overwrite Only
               zLevel = Asc(9)  ' Compression Level (0 - 9)
               zEncrypt = 0     ' Encryption = 1 For Password Else 0
               zComment = 0     ' Comment = 1 if required
            
               zArgc = 1           ' Number Of Elements Of mynames Array
               zZipFileName = txtFile.Text
               zZipFileNames.zFiles(0) = sRutaServidor & sFile
               zRootDir = ""    ' This Affects The Stored Path Name
              
               '-- Go Zip Them Up!
               retcode = VBZip32
               
               DeleteFile sRutaServidor & sFile
               imgVerde.Visible = True
               imgRojo.Visible = False
               Screen.MousePointer = vbDefault
               StatusBar.Panels(1).Text = "Terminado...."
               If retcode = 0 Then
                  MsgBox "El Backup se realizo satisfactoriamente", vbInformation, sMensaje
               Else
                  MsgBox "Hubo fallas en el empaquetado", vbCritical, sMensaje
               End If
               
          Case Is = 2  ' Browse
'               cdBrowse.DialogTitle = "Abrir Archivo"
'               cdBrowse.FileName = txtFile.Text
'               cdBrowse.ShowOpen
'               txtFile.Text = cdBrowse.FileName
'               sRuta = Mid(cdBrowse.FileName, 1, InStrRev(txtFile.Text, "\"))
'               cdBrowse.InitDir = sRuta
'               Open App.Path & "\RUTA.INI" For Output As #1
'               Print #1, sRuta
'               Close #1
               
               
               
               
   End Select
   Exit Sub
err:
   MsgBox err.Description & Chr(13) & "No se realizó ningun Backup", vbCritical, sMensaje
   Screen.MousePointer = vbDefault
   Exit Sub
End Sub

Private Sub Command1_Click()
  Isql = "exec BK_INFOREST '" & txtFile.Text & "','" & sMDB & "'"
'  Cn.Execute Isql
    StatusBar.Panels(1).Text = "Creado Backup...."
  Set RsBK = Lib.OpenRecordset(Isql, Cn)

'   If RsBK.RecordCount = 0 Then
    StatusBar.Panels(1).Text = "Terminado...."
      MsgBox "Se genero correctamente el Backup", vbInformation, sMensaje
       
'   End If
End Sub

Private Sub DtpFecha_Change()
   txtFile.Text = Mid(cdBrowse.FileName, 1, InStr(1, cdBrowse.FileName, "BK") - 1) & UCase("BK" & Format(dtpFecha.Value, "ddmmmyyyy")) & ".Zip"
End Sub

Private Sub Form_Load()
   imgRojo.Visible = False
   imgVerde.Visible = True
   Centrar Me
   dtpFecha.Value = Now()
   If FileExists(App.Path & "\RUTA.INI") Then
      Open App.Path & "\RUTA.INI" For Input As #1   ' Abre el archivo para recibir los datos.
      Do While Not EOF(1)                           ' Repite el bucle hasta el final del archivo.
         Input #1, sRuta                            ' Lee el carácter en dos variables
      Loop
      Close #1
      cdBrowse.InitDir = sRuta
   Else
      cdBrowse.InitDir = "C:\"
   End If
   
   'GCAA
   txtFile.Text = cdBrowse.InitDir ' & UCase("BK" & Format(dtpFecha.Value, "ddmmmyyyy")) & ".Zip"
   cdBrowse.FileName = txtFile.Text
   Screen.MousePointer = vbDefault
End Sub

Private Sub zip_ZipMinorStatus(ItemName As String, Percent As Long, Cancel As Long)
   If (Barra.StatusText <> ItemName) Then
      Barra.StatusText = ItemName
   End If
   
   If (Barra.StatusPercent <> Percent) Then
      Barra.StatusPercent = Percent
   End If
End Sub

