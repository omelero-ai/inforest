VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FrmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update"
   ClientHeight    =   7005
   ClientLeft      =   3465
   ClientTop       =   2775
   ClientWidth     =   10815
   Icon            =   "FrmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10815
   Begin MSComCtl2.Animation aniVideo 
      Height          =   540
      Left            =   45
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   953
      _Version        =   393216
      FullWidth       =   49
      FullHeight      =   36
   End
   Begin VB.PictureBox Picture 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   10755
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6345
      Width           =   10815
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
         Height          =   520
         Index           =   1
         Left            =   9600
         Picture         =   "FrmUpdate.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdOpcion 
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   520
         Index           =   0
         Left            =   8250
         Picture         =   "FrmUpdate.frx":0534
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Cuidado :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   15
         TabIndex        =   14
         Top             =   15
         Width           =   1230
      End
      Begin VB.Label Label3 
         Caption         =   "Verifique que todos los usuarios se encuentren fuera del sistema de Inforest para poder ejecutar este proceso."
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   1290
         TabIndex        =   13
         Top             =   60
         Width           =   4485
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5805
      Left            =   4230
      TabIndex        =   5
      Top             =   480
      Width           =   6555
      Begin TrueOleDBGrid80.TDBGrid grdGrilla 
         Height          =   5475
         Left            =   90
         TabIndex        =   7
         Top             =   225
         Width           =   6390
         _ExtentX        =   11271
         _ExtentY        =   9657
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   16
         Columns(0)._MaxComboItems=   5
         Columns(0).ValueItems(0)._DefaultItem=   0
         Columns(0).ValueItems(0).Value=   "0"
         Columns(0).ValueItems(0).Value.vt=   8
         Columns(0).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(0).ValueItems(0).DisplayValue(0)=   "bHQAADYDAABCTTYDAAAAAAAANgAAACgAAAAQAAAAEAAAAAEAGAAAAAAAAAMAAAAAAAAAAAAAAAAA"
         Columns(0).ValueItems(0).DisplayValue(1)=   "AAAAAAD///////////8An+IAneAAmt4Aldv///////8Ag8wAfsgAecQAdcD/////////////////"
         Columns(0).ValueItems(0).DisplayValue(2)=   "//8ApecAouVRxvJEvu8Amt7///////8AitFgxvI8quIAesQAdcH///////////////8AqOkzu++C"
         Columns(0).ValueItems(0).DisplayValue(3)=   "4P532vwAneEAm98Al9wkpuJq1/9n0PkZktQAe8X///////////////////8AqOo+wfGF4v86u+4A"
         Columns(0).ValueItems(0).DisplayValue(4)=   "nuEAm99NwvFs2P82ruUAhs7///////////////////////////8AqepIxfNy2PoAoeRSxvNz2/9B"
         Columns(0).ValueItems(0).DisplayValue(5)=   "u+0Ak9n///////////////////////////////8ArOwAqep93v1j1fxn2v9e0PkAnOAAmd7/////"
         Columns(0).ValueItems(0).DisplayValue(6)=   "//////////8At/YAtfQAs/IAsfAAru8ArO1d0vph2f9b1/9c0PoAn+IAneAAmt4AldsAj9YAidEA"
         Columns(0).ValueItems(0).DisplayValue(7)=   "uvh63/2k7P+B4P112/ty2vt54P9h2v9c2P9r2/9v1fpl0Ph11/uF3v9fy/UAkNYAvPkixPkAuPYA"
         Columns(0).ValueItems(0).DisplayValue(8)=   "tvUAtPMAsvFn2fxk2/9g2f9f0/sApecAo+UAoOMAneEaqOUAl9wAvvsAvPn///////8AtvUAtPOD"
         Columns(0).ValueItems(0).DisplayValue(9)=   "4v5w3v5x3f9m1fsAqOoApuj///////8AnuEAm9////////////////8AufdN0PqJ5f4AsvJY0PmB"
         Columns(0).ValueItems(0).DisplayValue(10)=   "4f9Gx/UAqer///////////////////////////8AvfpE0PuR6P9X0/oAtfQAs/JW0Pl+4P86wvQA"
         Columns(0).ValueItems(0).DisplayValue(11)=   "qer///////////////////8AwP0tzPyH5f6R6f8VwPgAt/YAtfQpwfWA4f9y2fwArO0Aquv/////"
         Columns(0).ValueItems(0).DisplayValue(12)=   "//////////8Awv4iyf1n3P554P4Au/n///////8AtvR63f1RzvkAr+8ArO3/////////////////"
         Columns(0).ValueItems(0).DisplayValue(13)=   "//8Awv4Awf0Av/wAvvv///////8AuPYAtvUAtPMAsvH/////////////////////////////////"
         Columns(0).ValueItems(0).DisplayValue(14)=   "//////////////////////////////////////////8="
         Columns(0).ValueItems(0).DisplayValue.vt=   9
         Columns(0).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(0).ValueItems(1)._DefaultItem=   0
         Columns(0).ValueItems(1).Value=   "1"
         Columns(0).ValueItems(1).Value.vt=   8
         Columns(0).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(0).ValueItems(1).DisplayValue(0)=   "bHQAADYDAABCTTYDAAAAAAAANgAAACgAAAAQAAAAEAAAAAEAGAAAAAAAAAMAAAAAAAAAAAAAAAAA"
         Columns(0).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
         Columns(0).ValueItems(1).DisplayValue(2)=   "//////////8AAPEAAPEAAPEAAPEAAO8AAO8AAO0AAO3///////////////////////////8AAPUa"
         Columns(0).ValueItems(1).DisplayValue(3)=   "IPU8TPk6Sfg4R/g1Rfg0Q/cyQvcUG/EAAO3///////////////////8AAPcdI/lEU/okKfkSEvcP"
         Columns(0).ValueItems(1).DisplayValue(4)=   "D/YMDPUJCfUWG/UzQ/cUG/EAAO3///////////8AAPkfJfpKWPtCR/vJyf07O/kTE/cQEPYzM/fF"
         Columns(0).ValueItems(1).DisplayValue(5)=   "xf0wNfc0RPcUG/IAAO3///////8AAPtPXf0yN/vLy/7y8v/r6/47O/k5Ofjq6v7x8f7Fxf0YHfYz"
         Columns(0).ValueItems(1).DisplayValue(6)=   "Q/cAAO////////8AAP1SX/0oKPxHR/zs7P/y8v/s7P/s7P7x8f/q6v40NPcLC/U1RfgAAO//////"
         Columns(0).ValueItems(1).DisplayValue(7)=   "//8AAP1VYv4sLP0pKfxISPzt7f/y8v/y8v/s7P46OvkSEvcPD/Y4SPgAAPH///////8AAP1XZP4w"
         Columns(0).ValueItems(1).DisplayValue(8)=   "MP0tLf1LS/zt7f/y8v/y8v/s7P89PfkWFvgTE/c8S/gAAPH///////8AAP9aZ/4zM/5QUP3t7f/z"
         Columns(0).ValueItems(1).DisplayValue(9)=   "8//t7f/t7f/y8v/s7P4+PvoXF/g/TvkAAPH///////8AAP9baP9DR/7Pz//z8//t7f9MTPxKSvzs"
         Columns(0).ValueItems(1).DisplayValue(10)=   "7P/y8v/Kyv4qL/pCUfoAAPP///////8AAP8mK/9dav9YW//Pz/9SUv4vL/0sLP1LS/zMzP5ITPtJ"
         Columns(0).ValueItems(1).DisplayValue(11)=   "V/sdI/kAAPX///////////8AAP8mK/9dav9DR/80NP4yMv4wMP0tLf04PPxPXfwfJfoAAPf/////"
         Columns(0).ValueItems(1).DisplayValue(12)=   "//////////////8AAP8mK/9caf9baP9aZ/5YZf5WY/5UYf4iJ/wAAPv/////////////////////"
         Columns(0).ValueItems(1).DisplayValue(13)=   "//////8AAP8AAP8AAP8AAP8AAP0AAP0AAP0AAP3/////////////////////////////////////"
         Columns(0).ValueItems(1).DisplayValue(14)=   "//////////////////////////////////////////8="
         Columns(0).ValueItems(1).DisplayValue.vt=   9
         Columns(0).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(0).ValueItems(2)._DefaultItem=   0
         Columns(0).ValueItems(2).Value=   "2"
         Columns(0).ValueItems(2).Value.vt=   8
         Columns(0).ValueItems(2).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
         Columns(0).ValueItems(2).DisplayValue(0)=   "bHQAADYDAABCTTYDAAAAAAAANgAAACgAAAAQAAAAEAAAAAEAGAAAAAAAAAMAAAAAAAAAAAAAAAAA"
         Columns(0).ValueItems(2).DisplayValue(1)=   "AAAAAAD/9P//7vz7/vXx//Hm/+nm/+3s//Lp//Hq//Po//Lk//Dj/+/t//Tz+vX/9///6v//+f//"
         Columns(0).ValueItems(2).DisplayValue(2)=   "/f72//Td/940ZDoqXDJBaks2W0E8YUclTzApXDYxYjzh/+nq/e7//P//+P/7//j0//LE8MdNilIy"
         Columns(0).ValueItems(2).DisplayValue(3)=   "eDtUml6Tz5+Zz6Si2K2NyZlIj1clbDMdWijk/+r0//fz+fTp/+rh/+RMjFEndzBftmye8axutXx7"
         Columns(0).ValueItems(2).DisplayValue(4)=   "vYp2tYOLz5iS4p9vwn0YZCQcVyTo/+3x//Lb/+FOhlU+jkllwHGX9qVqw3luuHxusn1xs39iqXBm"
         Columns(0).ValueItems(2).DisplayValue(5)=   "uXST7KJpu3M0ej09bUPc/9/d/+c8fEhXqGOa9alPrWBbsWuGypO388OAuYyAwIxWpWJhtW2E1o5O"
         Columns(0).ValueItems(2).DisplayValue(6)=   "k1YqWy/m/+rf/+tWj2KY4qZ90IxtwHyU3qTL/9bj/+vk/+2Rw5t2t4BmrG+U2Zyb1qMuVzLs/+7c"
         Columns(0).ValueItems(2).DisplayValue(7)=   "/+pVjWSS26N9zI10woXJ/9jW/+Li/+zo//Dh/+iGvY5+v4d3tn6b0qFIakvq/+zd/+1blWyQ26OG"
         Columns(0).ValueItems(2).DisplayValue(8)=   "15hgr3Cr77jE9c+HrY+84MLj/+rB/Mqe4qmGx4+e1qVLcFDk/+jO/91opXmH15p30It81I5suXt1"
         Columns(0).ValueItems(2).DisplayValue(9)=   "sYF7rYV8q4S/9MjT/92y/r6ByoyIxpIyXTjq//DW/+dssX5Zsm2L8aBv1YNoxnlntXWFyZBxtXxy"
         Columns(0).ValueItems(2).DisplayValue(10)=   "u32M45lrxnqv/71gq203bULj/+vY/+husH9ZsGxoznyO96BjynV1zIJdrGlmtHR4y4ZpyHiN8Z9l"
         Columns(0).ValueItems(2).DisplayValue(11)=   "wHUseTxTiGDP+9zd/+rY/+VptXtkv3Rw1YCQ9aB91IqT5J+C0o9+0YyZ+Khy04MuhUFHjVjf//Dm"
         Columns(0).ValueItems(2).DisplayValue(12)=   "//Xz//vq//TJ/9Ztt3tftW1nvnSO3ZqS252l7bGJ15dRp2FKn1tLlFzS/+Pq//ns/Pv47f33+v7s"
         Columns(0).ValueItems(2).DisplayValue(13)=   "//TU/95oqXJ4vIN3soBkm2xlnG1joG5YnGVho3DT/+To//jr8Pn/+P//9P//9P/0//nm/+7X/+Da"
         Columns(0).ValueItems(2).DisplayValue(14)=   "/+PV/93h/+ng/+jd/+fW/+La/+nh//Dp/fj/+P//7/8="
         Columns(0).ValueItems(2).DisplayValue.vt=   9
         Columns(0).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
         Columns(0).ValueItems.Count=   3
         Columns(0).Caption=   "Est"
         Columns(0).DataField=   "tEstado"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Descripción"
         Columns(1).DataField=   "tDescripcion"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Fecha"
         Columns(2).DataField=   "tFecha"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=741"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=661"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=1"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=6403"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=6324"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=3678"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=3598"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   2
         DefColWidth     =   0
         HeadLines       =   1.5
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         DirectionAfterTab=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         CollapseColor   =   128
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Named:id=33:Normal"
         _StyleDefs(43)  =   ":id=33,.parent=0"
         _StyleDefs(44)  =   "Named:id=34:Heading"
         _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(46)  =   ":id=34,.wraptext=-1"
         _StyleDefs(47)  =   "Named:id=35:Footing"
         _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(49)  =   "Named:id=36:Selected"
         _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(51)  =   "Named:id=37:Caption"
         _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(53)  =   "Named:id=38:HighlightRow"
         _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(55)  =   "Named:id=39:EvenRow"
         _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(57)  =   "Named:id=40:OddRow"
         _StyleDefs(58)  =   ":id=40,.parent=33"
         _StyleDefs(59)  =   "Named:id=41:RecordSelector"
         _StyleDefs(60)  =   ":id=41,.parent=34"
         _StyleDefs(61)  =   "Named:id=42:FilterBar"
         _StyleDefs(62)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Detalles del Proceso"
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   1740
         TabIndex        =   6
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5805
      Left            =   0
      TabIndex        =   0
      Top             =   495
      Width           =   4215
      Begin VB.Line Line1 
         BorderColor     =   &H00404000&
         X1              =   0
         X2              =   4200
         Y1              =   4860
         Y2              =   4860
      End
      Begin VB.Label lblBddLocal 
         AutoSize        =   -1  'True
         Caption         =   "Base Datos Central :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1605
         TabIndex        =   18
         Top             =   5370
         Width           =   1455
      End
      Begin VB.Label lblBddCentral 
         AutoSize        =   -1  'True
         Caption         =   "Base Datos Central :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1605
         TabIndex        =   17
         Top             =   5010
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Base Datos Local :"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   75
         TabIndex        =   16
         Top             =   5355
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Base Datos Central :"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   75
         TabIndex        =   15
         Top             =   4995
         Width           =   1455
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   2
         Left            =   120
         Picture         =   "FrmUpdate.frx":0ABE
         Top             =   1605
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   1
         Left            =   120
         Picture         =   "FrmUpdate.frx":0CD1
         Top             =   1125
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgProceso 
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "FrmUpdate.frx":0EE4
         Top             =   645
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblProceso 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finalizando la Actualización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   225
         Index           =   2
         Left            =   480
         TabIndex        =   4
         Top             =   1620
         Width           =   2475
      End
      Begin VB.Label lblProceso 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actualizando Tablas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   240
         Index           =   1
         Left            =   480
         TabIndex        =   3
         Top             =   1125
         Width           =   1845
      End
      Begin VB.Label lblProceso 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Verificando Conectividad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   645
         Width           =   2250
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Pasos para la Actualización"
         ForeColor       =   &H00404000&
         Height          =   195
         Left            =   555
         TabIndex        =   1
         Top             =   0
         Width           =   1950
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actualización de Datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   165
      TabIndex        =   11
      Top             =   90
      Width           =   2400
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   480
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "FrmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Conx As New Connection
Dim ConexServidor As New ADODB.Connection
Dim sDtemporal As String
Dim Parametro As String
Dim RsDtemporal  As Recordset
Dim RsTablas As Recordset
Dim RsTemp As Recordset

Sub Inicializa()
    lblProceso(1).ForeColor = &H808080
    lblProceso(2).ForeColor = &H808080
    
    imgProceso(1).Visible = False
    imgProceso(2).Visible = False
    
    
    Isql = "Delete From " & sDtemporal & " Where tTipo IN ('A', 'C')"
    Cn.Execute Isql
    Isql = "Update " & sDtemporal & " Set tFecha = '', tEstado = 0"
    Cn.Execute Isql
    TraeTemporal
    
    aniVideo.AutoPlay = True
    aniVideo.Visible = True
    DoEvents
End Sub

Sub TraeTemporal()
    Isql = "SELECT tEstado, tDescripcion, tFecha FROM " & sDtemporal
    Set RsDtemporal = Lib.OpenRecordset(Isql, Cn)
    Set grdGrilla.DataSource = RsDtemporal
End Sub

Function VerificaConexion() As Boolean
    If sServidorCentral <> "" Then
        If sServidorCentral <> sRuta Or bdInforestCentral <> sMDB Then
            Set Conx = New Connection
            Conx.Provider = "SQLOLEDB"
            Conx.CursorLocation = adUseServer
            Conx.ConnectionString = "User ID=" & sUserName & _
                                    ";password=" & sUserPassword & _
                                    ";Data Source=" & sServidorCentral & _
                                    ";Initial Catalog=" & bdInforestCentral
            Conx.ConnectionTimeout = 5
            Conx.Open
            If Conx.State Then
                VerificaConexion = True
            Else
                VerificaConexion = False
            End If
        End If
    End If
End Function

Private Sub cmdOpcion_Click(Index As Integer)
On Error GoTo LeerError

    If Index = 0 Then
        If VerificaConexion Then
         If MsgBox("¿Seguro que desea continuar con la actualización de los datos?", vbQuestion + vbYesNo, sMensaje) = vbYes Then
         
           Screen.MousePointer = vbHourglass
'
           Inicializa
'
           Isql = "UPDATE " & sDtemporal & " SET tEstado = 2, tFecha = '" & Format(Now, "dd/MM/yyyy hh:mm:ss am/pm") & "' Where tTipo = 'I'"
            Cn.Execute Isql
            imgProceso(0).Visible = True
            DoEvents
            Sleep (1500)
'
            'Actualizacion de Tablas
            lblProceso(1).ForeColor = &H80000012
            Isql = "Select 0, tAcceso.tDescripcion From tAcceso, tAccesoEnvia Where tacceso.tCodigoAcceso = tAccesoEnvia.tCodigoAcceso "
            'Debug.Print Isql
            Set RsTablas = Lib.OpenRecordset(Isql, Cn)
            While Not RsTablas.EOF
                Isql = "INSERT INTO " & sDtemporal & " VALUES(0, '" & RsTablas!tDescripcion & "', '', 'A')"
                Cn.Execute Isql

                RsTablas.MoveNext
            Wend
            TraeTemporal
            DoEvents
            Sleep (500)
            'Actualiza valores
            Parametro = ""
            Set RsTemp = Lib.OpenRecordset("SELECT dbo.TACCESO.tCodigoAcceso, ISNULL(dbo.TACCESO.nOrden, 0) AS norden, ISNULL(dbo.TACCESOENVIA.lEnvia, 0) AS lenvia FROM dbo.TACCESO LEFT OUTER JOIN dbo.TACCESOENVIA ON dbo.TACCESO.tCodigoAcceso = dbo.TACCESOENVIA.tCodigoAcceso Where (IsNull(dbo.TACCESO.nOrden, 0) <> 0)  ORDER BY ISNULL(dbo.TACCESO.nOrden, 0)", Cn)
            If Not (RsTemp.EOF Or RsTemp.BOF) Then
                         RsTemp.MoveFirst
                         For X = 0 To RsTemp.RecordCount - 1
                             Parametro = Parametro & Trim(str(IIf(RsTemp.Fields("lenvia") = False, "0", "1")))
                               RsTemp.MoveNext
                         Next X
                        Cn.Execute "sp_ActualizaTablas '" & sServidorCentral & "','" & bdInforestCentral & "','" & Parametro & "'"
          


                Isql = "UPDATE " & sDtemporal & " SET tEstado = 2, tFecha = '" & Format(Now, "dd/MM/yyyy hh:mm:ss am/pm") & "' Where tTipo = 'A'"
                Cn.Execute Isql
                TraeTemporal
                imgProceso(1).Visible = True
            End If
'

                DoEvents
                Sleep (1500)
'
                lblProceso(2).ForeColor = &H80000012
                imgProceso(2).Visible = True
'            End If
'
            DoEvents
            aniVideo.AutoPlay = False
            aniVideo.Visible = False

        Else
            
            Isql = "UPDATE " & sDtemporal & " SET tEstado = 1"
            Cn.Execute Isql
            TraeTemporal
        End If
        End If
        Screen.MousePointer = vbDefault
    Else
        Unload Me
    End If
    Exit Sub

LeerError:
    DoEvents
    aniVideo.AutoPlay = False
    aniVideo.Visible = False

    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, sMensaje
End Sub

Private Sub Form_Load()
    Dim nFrames As Long

    Centrar Me
    
    lblBddCentral.Caption = UCase(bdInforestCentral)
    lblBddLocal.Caption = UCase(sMDB)
    sDtemporal = dbTemporal(sUsuario, 4, "tEstado", "int", _
                                         "tDescripcion", "nVarChar(100)", _
                                         "tFecha", "Varchar(30)", _
                                         "tTipo", "char(1)")
    Isql = "INSERT INTO " & sDtemporal & " VALUES(0, 'IP Servidor Central : " & sServidorCentral & "', '', 'I')"
    Cn.Execute Isql
    On Error Resume Next
    aniVideo.Open App.Path & "\bmps\FileMove.avi"
    TraeTemporal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Isql = "DELETE FROM " & sDtemporal
    Cn.Execute Isql
End Sub

