VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUserSenhaInc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alteração de Senha"
   ClientHeight    =   3615
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   5940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   5940
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2550
      Width           =   5940
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "ESC"
         Height          =   880
         Left            =   3030
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   90
         Width           =   1215
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "ENTER"
         Default         =   -1  'True
         Height          =   880
         Left            =   1830
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   90
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   2340
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4128
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   529
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Dados"
      TabPicture(0)   =   "userSenhaInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraProf"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraProf 
         Height          =   1695
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   5295
         Begin VB.TextBox txtConfSenha 
            Height          =   288
            IMEMode         =   3  'DISABLE
            Left            =   1560
            MaxLength       =   6
            PasswordChar    =   "#"
            TabIndex        =   1
            Text            =   "txtCon"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtNovaSenha 
            Height          =   288
            IMEMode         =   3  'DISABLE
            Left            =   1560
            MaxLength       =   6
            PasswordChar    =   "#"
            TabIndex        =   0
            Text            =   "txtNov"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtUsuario 
            BackColor       =   &H00E0E0E0&
            Height          =   288
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   4
            TabStop         =   0   'False
            Text            =   "txtUsuario"
            Top             =   360
            Width           =   3255
         End
         Begin VB.Label Label3 
            Caption         =   "Confirmar Senha"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Nova Senha"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Usuário"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "frmUserSenhaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public lngCONTROLEACESSOID  As Long
Public strUsuario           As String



Private Sub cmdCancelar_Click()
On Error GoTo trata
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub


Private Sub cmdConfirmar_Click()
  On Error GoTo trata
  Dim objUsuario As busSisMed.clsUsuario
  '
  If Not ValidaCampos Then Exit Sub
  '
  Set objUsuario = New busSisMed.clsUsuario
  objUsuario.AlterarSenha lngCONTROLEACESSOID, _
                          Encripta(UCase$(txtNovaSenha.Text))
  '
  MsgBox "A senha do usuário " & txtUsuario.Text & _
     " foi alterada com sucesso!", vbExclamation, TITULOSISTEMA
  '
  Set objUsuario = Nothing
  Unload Me
  '
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub Form_Activate()
  SetarFoco txtNovaSenha
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 39 Then
    KeyAscii = 96
  End If
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  AmpS
  'Me.Caption = Me.Caption & sgFormCaption
  Me.Width = 6060
  Me.Height = 3945
  '
  CenterForm Me
  LerFiguras Me, tpBmp_Vazio, cmdConfirmar, cmdCancelar
  '
  txtUsuario.Text = strUsuario
  txtNovaSenha.Text = ""
  txtConfSenha.Text = ""
  '
  AmpN
  '
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub
Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg        As String
  Dim blnSetarFoco  As Boolean
  '
  blnSetarFoco = True
  If Len(txtNovaSenha.Text) < 4 Then
    strMsg = strMsg & "Informar a nova Senha com mínimo de 4 dígitos" & vbCrLf
    blnSetarFoco = False
    Pintar_Controle txtNovaSenha, tpCorContr_Erro
    SetarFoco txtNovaSenha
  End If
  '
  If txtNovaSenha.Text <> txtConfSenha Then
    strMsg = strMsg & "Senhas digitadas não conferem" & vbCrLf
    Pintar_Controle txtConfSenha, tpCorContr_Erro
    If blnSetarFoco Then
      blnSetarFoco = False
      SetarFoco txtConfSenha
    End If
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  '
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  ValidaCampos = False
End Function

Private Sub txtConfSenha_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 8 Then
    If Len(txtConfSenha.Text) = 6 Then
      MsgBox "Asenha deve ter no máximo 6 palavras", vbExclamation, TITULOSISTEMA
    End If
  End If
End Sub

Private Sub txtConfSenha_LostFocus()
  Pintar_Controle txtConfSenha, tpCorContr_Normal
End Sub

Private Sub txtConfSenha_Gotfocus()
  Seleciona_Conteudo_Controle txtConfSenha
End Sub

Private Sub txtNovaSenha_KeyPress(KeyAscii As Integer)
  If KeyAscii <> 8 Then
    If Len(txtNovaSenha.Text) = 6 Then
      MsgBox "Asenha deve ter no máximo 6 dígitos", vbExclamation, TITULOSISTEMA
    End If
  End If
End Sub

Private Sub txtNovaSenha_LostFocus()
  Pintar_Controle txtNovaSenha, tpCorContr_Normal
End Sub
Private Sub txtNovaSenha_Gotfocus()
  Seleciona_Conteudo_Controle txtNovaSenha
End Sub

