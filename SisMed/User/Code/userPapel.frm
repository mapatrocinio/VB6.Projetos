VERSION 5.00
Begin VB.Form frmUserPapel 
   Caption         =   "Imagem do Papel de Parede"
   ClientHeight    =   5250
   ClientLeft      =   1530
   ClientTop       =   1395
   ClientWidth     =   6210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5250
   ScaleWidth      =   6210
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5895
      Begin VB.DirListBox Dir1 
         ForeColor       =   &H00000000&
         Height          =   2115
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   2895
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
      Begin VB.FileListBox File1 
         ForeColor       =   &H00000000&
         Height          =   1065
         Left            =   3360
         Pattern         =   "*.bmp;*.jpg;*.gif"
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Diretório:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Drive:"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   2055
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Previssão -->"
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   3600
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdCancela 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   880
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ENTER"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   880
      Left            =   1860
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
   End
End
Attribute VB_Name = "frmUserPapel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancela_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim strPathArquivo As String
  '
  AmpS
  '
  If File1.FileName = "" Then
    AmpN
    Exit Sub
  End If
  If Right$(Dir1.Path, 1) = "\" Then
    strPathArquivo = Dir1.Path
  Else
    strPathArquivo = Dir1.Path & "\"
  End If
  '
  gsBMP = strPathArquivo & File1.FileName
  CapturaParametrosRegistro 2
  '
  If Len(Trim(gsBMP)) <> 0 Then
    frmMDI.Picture = LoadPicture(gsBMP)
  End If
  MsgBox "Papel de parede adicionado com sucesso!", vbExclamation, TITULOSISTEMA
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
  
End Sub

Private Sub Dir1_Change()
  On Error GoTo trata
  File1.Path = Dir1.Path
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub Drive1_Change()
  On Error GoTo trata
  Dir1.Path = Drive1.Drive
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub


Private Sub File1_Click()
  On Error GoTo trata
  Dim cPathArquivo As String
  AmpS
  '
  If Right$(Dir1.Path, 1) = "\" Then
    cPathArquivo = Dir1.Path
  Else
    cPathArquivo = Dir1.Path & "\"
  End If
  Image1.Picture = LoadPicture(cPathArquivo & File1.FileName)
  Image1.Refresh
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  AmpS
  CenterForm Me
  Dir1.Path = gsBMPPath
  Drive1.Drive = "C:\"
  LerFiguras Me, tpBmp_Vazio, Me.cmdOk, , Me.cmdCancela
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub


