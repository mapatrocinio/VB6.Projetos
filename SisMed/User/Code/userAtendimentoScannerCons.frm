VERSION 5.00
Begin VB.Form frmUserAtendimentoScannerCons 
   Caption         =   "Imagens Scanner"
   ClientHeight    =   8490
   ClientLeft      =   1530
   ClientTop       =   1395
   ClientWidth     =   11880
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
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   8490
      Left            =   10170
      ScaleHeight     =   8490
      ScaleWidth      =   1710
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1710
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2025
         Left            =   0
         ScaleHeight     =   1965
         ScaleWidth      =   1575
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   6390
         Width           =   1635
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8415
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   10065
      Begin VB.PictureBox Picture1 
         Height          =   8175
         Left            =   3030
         ScaleHeight     =   8115
         ScaleWidth      =   6885
         TabIndex        =   6
         Top             =   150
         Width           =   6945
         Begin VB.Image Image1 
            Height          =   8085
            Left            =   30
            Stretch         =   -1  'True
            Top             =   30
            Width           =   6855
         End
      End
      Begin VB.DirListBox Dir1 
         ForeColor       =   &H00000000&
         Height          =   3915
         Left            =   90
         TabIndex        =   2
         Top             =   540
         Width           =   2895
      End
      Begin VB.FileListBox File1 
         ForeColor       =   &H00000000&
         Height          =   3405
         Left            =   90
         Pattern         =   "*.bmp;*.jpg;*.gif"
         TabIndex        =   1
         Top             =   4800
         Width           =   2925
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Arquivo:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   90
         TabIndex        =   7
         Top             =   4500
         Width           =   855
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Diretório:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   90
         TabIndex        =   3
         Top             =   180
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmUserAtendimentoScannerCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strCaminhoFinal As String
Public strArquivoFinal As String

Private Sub cmdCancelar_Click()
  Unload Me
End Sub



Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim strPathArquivo As String
  '
  AmpS
  '
  If File1.FileName = "" Then
    strPathFinal = ""
    TratarErroPrevisto "Selecione uma receita", "[frmUserAtendimentoScannerCons.cmdConfirmar_Click]"
    AmpN
    Exit Sub
  End If
  If Right$(Dir1.Path, 1) = "\" Then
    strPathArquivo = Dir1.Path
  Else
    strPathArquivo = Dir1.Path & "\"
  End If
  '
  strCaminhoFinal = strPathArquivo
  strArquivoFinal = File1.FileName
  AmpN
  Unload Me
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

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco File1
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserAtendimentoScannerCons.Form_Activate]"
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim objGeral                As busSisMed.clsGeral
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objAtendimento          As busSisMed.clsAtendimento
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 9000
  Me.Width = 12000
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  strCaminhoFinal = ""
  strArquivoFinal = ""
  Dir1.Path = gsPathLocal
  LerFiguras Me, tpBmp_Vazio, Me.cmdOk, , Me.cmdCancelar
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub
