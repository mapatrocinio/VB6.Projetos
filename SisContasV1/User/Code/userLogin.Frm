VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2790
   ClientLeft      =   4320
   ClientTop       =   3675
   ClientWidth     =   3330
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2790
   ScaleWidth      =   3330
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Default         =   -1  'True
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtSenha 
      Height          =   312
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1404
      Width           =   1452
   End
   Begin VB.TextBox txtUsuario 
      Height          =   312
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   1
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
      Height          =   252
      Left            =   576
      TabIndex        =   2
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
      TabIndex        =   0
      Top             =   1008
      Width           =   792
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public QuemChamou As Integer
'QuemChamou Assume 0 Fora do Sistema 1 Dentro do sistema
Option Explicit

Private Sub cmdCancelar_Click()
  Select Case QuemChamou
  Case 0
    End
  Case Else
    Unload Me
  End Select
End Sub


Private Sub cmdOK_Click()
Dim sAux As String
Dim iRet As Integer
   On Error GoTo Erro_cmdOk_Click
   Screen.MousePointer = vbHourglass
   ''MsgStatus "Inicializando o Sistema e verificando o usuário . . ."
   
   'CarregaParametros
   sAux = Trim$(txtUsuario.Text)
   
   If Len(sAux) = 0 Then
      MsgBox "Entre com seu nome e senha.", vbOKOnly, TITULOSISTEMA
      Screen.MousePointer = vbDefault
      txtUsuario.SetFocus
      Exit Sub
   End If
      
   If Not ValidaUsuario() Then
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
            
   'frmMDI.statusbar1.Panels(1) = _
          UCase$(Left$(sAux, 1)) & LCase$(Right$(sAux, Len(sAux) - 1))
      
   ''MsgStatusPronto
   Screen.MousePointer = vbDefault
   'frmSplash.Timer2.Interval = 2000
    If QuemChamou = 0 Then
      frmSplash.Timer2.Interval = 1000
      frmSplash.chkSistema.Value = 1
    End If
    frmMDI.mnuArquivo(0).Caption = "&Desconectar"
    '
    'Monta_Menu 1
   
Fim_cmdOk_Click:
   Unload Me
   Exit Sub
   
Erro_cmdOk_Click:
   Select Case Err
   Case 3024
      MsgBox "O arquivo de dados não foi encontrado. Verifique a configuração do arquivo de inicialização ou do registry." _
          & Chr(10) & "Saindo do sistema.", 16, TITULOSISTEMA
      End
   Case Else
      Call MsgBox("O seguinte erro ocorreu: " & Err.Number & " - " & Err.Description, vbOKOnly, TITULOSISTEMA)
   End Select


End Sub

Function ValidaUsuario() As Boolean
   'Esta rotina vai validar a existência do usuário no cadastro
   'e checar sua senha.
   
   'Saídas: true - OK
   '            false - Erro
   
   Dim iRet As Integer, ICont As Integer
   Dim strSql As String
   Dim ssAcesso As ADODB.Recordset
   Dim objGeral As busSisContas.clsGeral
  
   Set objGeral = New busSisContas.clsGeral
   
   On Error GoTo Erro_ValidaUsuario

   strSql = "Select usuario, senha, nivel "
   strSql = strSql & " FROM ControleAcesso "
   strSql = strSql & " WHERE usuario =  '" & Trim(txtUsuario.Text) & "';"

   Set ssAcesso = objGeral.ExecutarSQL(strSql)
   'Verifica se o usuário existe
   If ssAcesso.EOF Then
    Err = 32000
    GoTo Erro_ValidaUsuario
   End If
    '
   ssAcesso.MoveFirst
   
   'Checa a senha do usuário
   If UCase$(Encripta(ssAcesso.Fields("Senha").Value & "")) <> UCase$(txtSenha.Text) Then
    Err = 31998
    GoTo Erro_ValidaUsuario
   End If
   '
   'Captura configurações do Usuário
   gsNomeUsu = UCase$(txtUsuario.Text)
   gsNivel = ssAcesso!nivel & ""
   'gsSecao = ssAcesso!Secao
   '
   ssAcesso.Close
   Set ssAcesso = Nothing
   '
   CapturaParametrosRegistro 1
   '
   ValidaUsuario = True
   Set objGeral = Nothing
Fim_ValidaUsuario:
   Exit Function
            
Erro_ValidaUsuario:
   Set objGeral = Nothing
   Select Case Err
      Case 3044
         MsgBox "O caminho especificado para os arquivos do sistema não foi encontrado." & Chr(10) & "O programa será terminado.", _
            16, "Erro de Abertura da Base de Consultas"
         End
      Case 32000     'Usuário não existente
         MsgBox "Usuário não encontrado.", vbExclamation, TITULOSISTEMA
         txtUsuario.SelStart = 0
         txtUsuario.SelLength = Len(txtUsuario.Text)
         txtUsuario.SetFocus
      Case 31999     'Usuário não tem permissão neste sistema
         MsgBox "Usuário não tem permissão neste sistema.", vbExclamation, TITULOSISTEMA
      Case 31998     'Senha errada
         MsgBox "Senha inválida.", vbExclamation, TITULOSISTEMA
         txtSenha.SelStart = 0
         txtSenha.SelLength = Len(txtUsuario.Text)
         txtSenha.SetFocus
      Case Else
         Call MsgBox("O seguinte erro ocorreu: " & Err.Number & " - " & Err.Description, vbExclamation, TITULOSISTEMA)
   End Select
   ValidaUsuario = False
   GoTo Fim_ValidaUsuario
      
End Function

Private Sub Form_Activate()
  Screen.MousePointer = vbDefault
  lblTitulo.Caption = TITULOSISTEMA & " Vs. " & App.Major & "." & App.Minor
  'lblVersao.Caption = lblVersao.Caption & App.Major & "." & App.Minor
  '
  If Len(Trim(gsNomeUsu)) = 0 Then
    txtUsuario.SetFocus
  Else
    txtUsuario.Text = gsNomeUsu
    txtSenha.SetFocus
  End If
End Sub

Private Sub Form_Load()
   AmpS
   CenterForm Me
   LerFiguras Me, tpBmp_Login, Me.cmdOk, Me.cmdCancelar
   
   AmpN
End Sub

Private Sub txtSenha_GotFocus()
    txtSenha.SelStart = 0
    txtSenha.SelLength = Len(txtSenha.Text & "")
End Sub

Private Sub txtUsuario_GotFocus()
    txtUsuario.SelStart = 0
    txtUsuario.SelLength = Len(txtUsuario.Text & "")
End Sub

