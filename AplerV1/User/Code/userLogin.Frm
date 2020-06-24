VERSION 5.00
Begin VB.Form frmUserLogin 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2790
   ClientLeft      =   4320
   ClientTop       =   3675
   ClientWidth     =   3330
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   Icon            =   "userLogin.frx":0000
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
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1830
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   880
      Left            =   1830
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1830
      Width           =   1215
   End
   Begin VB.TextBox txtSenha 
      Height          =   312
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1404
      Width           =   1452
   End
   Begin VB.TextBox txtUsuario 
      Height          =   312
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1020
      Width           =   1452
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
      TabIndex        =   6
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
      Height          =   252
      Left            =   576
      TabIndex        =   5
      Top             =   1404
      Width           =   792
   End
   Begin VB.Label lblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuário"
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
      Height          =   252
      Left            =   576
      TabIndex        =   4
      Top             =   1008
      Width           =   792
   End
End
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public QuemChamou As Integer
'QuemChamou Assume 0 Fora do Sistema 1 Dentro do sistema
Option Explicit

Private Sub cmdCancelar_Click()
  On Error GoTo trata
  Select Case QuemChamou
  Case 0
    End
  Case Else
    Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  End
End Sub


Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim blnSetarFocoControle As Boolean
  blnSetarFocoControle = True
  AmpS
  If Not Valida_String(txtUsuario, TpObrigatorio, blnSetarFocoControle) Then
    TratarErroPrevisto "Informe o usuário", _
                       "[ValidaUsuario]"
    Exit Sub
  End If
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
  If QuemChamou = 0 Then
    frmUserSplash.Timer2.Interval = 1000
    frmUserSplash.chkSistema.Value = 1
  End If
  frmMDI.mnuArquivo(0).Caption = "&Desconectar"
  '
  Monta_Menu 0

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
  Dim objGeral As busApler.clsGeral
  '
  Set objGeral = New busApler.clsGeral
  '
  strSql = "Select usuario, senha, nivel "
  strSql = strSql & " FROM ControleAcesso "
  strSql = strSql & " WHERE usuario =  " & Formata_Dados(txtUsuario.Text, tpDados_Texto, tpNulo_Aceita)

  Set objRs = objGeral.ExecutarSQL(strSql)
  'Verifica se o usuário existe
  If objRs.EOF Then
    TratarErroPrevisto "Usuário não encontrado", _
                       "[ValidaUsuario]"
    Pintar_Controle txtUsuario, tpCorContr_Erro
    SetarFoco txtUsuario
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Exit Function
  End If
  'Checa a senha do usuário
  If UCase$(Encripta(objRs.Fields("Senha").Value)) <> UCase$(txtSenha.Text) Then
    TratarErroPrevisto "Senha inválida", _
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
  gsNomeUsu = UCase$(txtUsuario.Text)
  gsNivel = Trim(objRs.Fields("nivel").Value & "")
  '
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  '
  CapturaParametrosRegistro 1
  '
  ValidaUsuario = True

  Exit Function
            
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  lblTitulo.Caption = TITULOSISTEMA & " Vs. " & App.Major & "." & App.Minor
  If Len(Trim(gsNomeUsu)) = 0 Then
    SetarFoco txtUsuario
  Else
    txtUsuario.Text = gsNomeUsu
    SetarFoco txtSenha
  End If
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

Private Sub txtUsuario_GotFocus()
  Seleciona_Conteudo_Controle txtUsuario
End Sub

Private Sub txtUsuario_LostFocus()
  Pintar_Controle txtUsuario, tpCorContr_Normal
End Sub
