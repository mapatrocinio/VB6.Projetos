VERSION 5.00
Begin VB.Form frmUserLoginSup 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2790
   ClientLeft      =   4320
   ClientTop       =   3675
   ClientWidth     =   3330
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   Icon            =   "userLoginSup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2790
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "ENTER"
      Default         =   -1  'True
      Height          =   880
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1650
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   880
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1650
      Width           =   1215
   End
   Begin VB.TextBox txtSenha 
      Height          =   312
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1170
      Width           =   1452
   End
   Begin VB.Label lblVersao 
      BackStyle       =   0  'Transparent
      Caption         =   "Versão : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   90
      TabIndex        =   4
      Top             =   120
      Width           =   3195
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSenha 
      BackStyle       =   0  'Transparent
      Caption         =   "&Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   570
      TabIndex        =   3
      Top             =   1170
      Width           =   795
   End
End
Attribute VB_Name = "frmUserLoginSup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public QuemChamou As Integer
'QuemChamou Assume 0 Fora do Sistema 1 Dentro do sistema
Option Explicit

Private Sub cmdCancelar_Click()
  Unload Me
End Sub


Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim blnSetarFocoControle As Boolean
  blnSetarFocoControle = True
  AmpS
  If Not Valida_String(txtSenha, TpObrigatorio, blnSetarFocoControle) Then
    TratarErroPrevisto "Informe a senha", _
                       "[ValidaUsuario]"
    Exit Sub
  End If
  
  If Not ValidaUsuario() Then
    AmpN
    Exit Sub
  End If
  
  AmpN
  Unload Me
  Exit Sub
   
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

'Esta rotina vai validar a existência do usuário no cadastro
'e checar sua senha.
'Saídas: true - OK
'            false - Erro
Function ValidaUsuario() As Boolean
  On Error GoTo trata
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  Dim objGeral As busSisLoc.clsGeral
  '
  Set objGeral = New busSisLoc.clsGeral
  '
  strSql = "Select FUNCIONARIO.usuario, FUNCIONARIO.senha, FUNCIONARIO.nivel, FUNCIONARIO.PRONTUARIOID, PRONTUARIO.NOME "
  strSql = strSql & " FROM FUNCIONARIO INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = FUNCIONARIO.PRONTUARIOID "
  strSql = strSql & " WHERE FUNCIONARIO.Senha =  " & Formata_Dados(Encripta(UCase$(txtSenha.Text)), tpDados_Texto, tpNulo_Aceita)
  strSql = strSql & " AND FUNCIONARIO.NIVEL IN('DIR', 'GER', 'ADM') "
  strSql = strSql & " AND INDEXCLUIDO =  " & Formata_Dados("N", tpDados_Texto, tpNulo_Aceita)

  Set objRs = objGeral.ExecutarSQL(strSql)
  'Verifica se o usuário existe
  If objRs.EOF Then
    TratarErroPrevisto "Senha não encontrada", _
                       "[ValidaUsuario]"
    Pintar_Controle txtSenha, tpCorContr_Erro
    SetarFoco txtSenha
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Exit Function
  End If
  'Checa a senha do usuário por nível
  If Not ((UCase$(objRs.Fields("nivel").Value & "") = gsAdmin) Or _
          (UCase$(objRs.Fields("nivel").Value & "") = gsDiretor) Or _
          (UCase$(objRs.Fields("nivel").Value & "") = gsGerente)) Then
    TratarErroPrevisto "Está senha não tem permissão de acesso a esta função do sistema", _
                       "[ValidaUsuario]"
    Pintar_Controle txtSenha, tpCorContr_Erro
    SetarFoco txtSenha
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Exit Function
  End If
  '
  'Captura configurações do Usuário
  gsNomeUsuLib = objRs.Fields("usuario").Value & ""
  gsNivelUsuLib = objRs.Fields("nivel").Value & ""
  Unload Me
  Exit Function
   
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  lblTitulo.Caption = TITULOSISTEMA & " Vs. " & App.Major & "." & App.Minor
  SetarFoco txtSenha
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  AmpS
  CenterForm Me
  LerFiguras Me, tpBmp_Login, Me.cmdOk, Me.cmdCancelar
  
  gsNomeUsuLib = ""
  gsNivelUsuLib = ""
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub
Private Sub txtSenha_GotFocus()
  Seleciona_Conteudo_Controle txtSenha
End Sub

Private Sub txtSenha_LostFocus()
  Pintar_Controle txtSenha, tpCorContr_Normal
End Sub


