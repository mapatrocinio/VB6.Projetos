VERSION 5.00
Begin VB.Form frmUserLogin 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6225
   ClientLeft      =   4320
   ClientTop       =   3675
   ClientWidth     =   3630
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   Icon            =   "userLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6225
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRegBMP 
      Height          =   312
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   12
      Top             =   5250
      Width           =   2085
   End
   Begin VB.TextBox txtRegEmpresa 
      Height          =   312
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   11
      Top             =   4920
      Width           =   2085
   End
   Begin VB.TextBox txtRegPathRpt 
      Height          =   312
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   10
      Top             =   4590
      Width           =   2085
   End
   Begin VB.TextBox txtRegPathIcons 
      Height          =   312
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   9
      Top             =   4260
      Width           =   2085
   End
   Begin VB.TextBox txtRegPathBmp 
      Height          =   312
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   8
      Top             =   3930
      Width           =   2085
   End
   Begin VB.TextBox txtRegPathBackup 
      Height          =   312
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   7
      Top             =   3600
      Width           =   2085
   End
   Begin VB.TextBox txtRegPathApp 
      Height          =   312
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   6
      Top             =   3270
      Width           =   2085
   End
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "Registrar"
      Height          =   345
      Left            =   2520
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5730
      Width           =   1035
   End
   Begin VB.TextBox txtRegBd 
      Height          =   312
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2940
      Width           =   2085
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   ">>"
      Height          =   345
      Left            =   3180
      TabIndex        =   4
      Top             =   2370
      Width           =   345
   End
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
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome bmp"
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
      Left            =   270
      TabIndex        =   24
      Top             =   5310
      Width           =   1155
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
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
      Left            =   270
      TabIndex        =   23
      Top             =   4980
      Width           =   1155
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Path rpt"
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
      Left            =   270
      TabIndex        =   22
      Top             =   4650
      Width           =   1155
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Path icons"
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
      Left            =   270
      TabIndex        =   21
      Top             =   4320
      Width           =   1155
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Path bmp"
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
      Left            =   270
      TabIndex        =   20
      Top             =   3990
      Width           =   1155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Path Backup"
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
      Left            =   270
      TabIndex        =   19
      Top             =   3660
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Path Aplic."
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
      Left            =   270
      TabIndex        =   18
      Top             =   3330
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Servidor BD"
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
      Left            =   270
      TabIndex        =   17
      Top             =   3000
      Width           =   1155
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
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
Public intHeightNormal As Integer
Public intHeightMax As Integer

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
    SetarFoco txtUsuario
    txtUsuario_GotFocus
    Exit Sub
  End If
  If Not Valida_String(txtSenha, TpObrigatorio, blnSetarFocoControle) Then
    TratarErroPrevisto "Informe a senha", _
                       "[ValidaUsuario]"
    SetarFoco txtSenha
    txtSenha_GotFocus
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
  Dim strSql        As String
  Dim objRs         As ADODB.Recordset
  Dim objGeral      As busSisMetal.clsGeral
  Dim objRegister   As busSisMetal.clsRegistro
  '
  Set objGeral = New busSisMetal.clsGeral
  '
  strSql = "Select FUNCIONARIO.usuario, FUNCIONARIO.senha, FUNCIONARIO.nivel, FUNCIONARIO.PESSOAID, PESSOA.NOME "
  strSql = strSql & " FROM FUNCIONARIO INNER JOIN PESSOA ON PESSOA.PKID = FUNCIONARIO.PESSOAID "
  strSql = strSql & " WHERE usuario =  " & Formata_Dados(txtUsuario.Text, tpDados_Texto, tpNulo_Aceita)
  strSql = strSql & " AND INDEXCLUIDO =  " & Formata_Dados("N", tpDados_Texto, tpNulo_Aceita)

  Set objRs = objGeral.ExecutarSQL(strSql)
  'Verifica se o usuário existe
  If objRs.EOF Then
    TratarErroPrevisto "Usuário não encontrado", _
                       "[ValidaUsuario]"
    Pintar_Controle txtUsuario, tpCorContr_Erro
    SetarFoco txtUsuario
    txtUsuario_GotFocus
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
    txtSenha_GotFocus
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Exit Function
  End If
  '
  'Captura configurações do Usuário
  gsNomeUsu = UCase$(txtUsuario.Text)
  gsNivel = Trim(objRs.Fields("nivel").Value & "")
  giFuncionarioId = objRs.Fields("PESSOAID").Value
  gsNomeUsuCompleto = objRs.Fields("NOME").Value
  '
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  '
  Set objRegister = New busSisMetal.clsRegistro
  objRegister.SalvaParametrosRegisterUsuario TITULOSISTEMA, _
                                             gsNomeUsu, _
                                             gsNivel
  'CapturaParametrosRegistro 1
  Set objRegister = Nothing
  '
  ValidaUsuario = True

  Exit Function
            
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Function

Private Sub cmdRegister_Click()
  On Error GoTo trata
  '
  If cmdRegister.Caption = ">>" Then
    cmdRegister.Caption = "<<"
    Me.Height = intHeightMax
    txtRegBd.TabStop = True
    SetarFoco txtRegBd
  Else
    cmdRegister.Caption = ">>"
    Me.Height = intHeightNormal
    txtRegBd.TabStop = False
    If Len(txtUsuario) = 0 Then
      SetarFoco txtUsuario
    Else
      SetarFoco txtSenha
    End If
  End If
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub


Private Sub cmdRegistrar_Click()
  On Error GoTo trata
  Dim objRegister As busSisMetal.clsRegistro
  'Salva informação no register
  Set objRegister = New busSisMetal.clsRegistro
  objRegister.SalvarChaveRegistro TITULOSISTEMA, _
                                  "ServidorBD", _
                                  txtRegBd.Text
  objRegister.SalvarChaveRegistro TITULOSISTEMA, _
                                  "CaminhoApp", _
                                  txtRegPathApp.Text
  objRegister.SalvarChaveRegistro TITULOSISTEMA, _
                                  "CaminhoBackUp", _
                                  txtRegPathBackup.Text
  objRegister.SalvarChaveRegistro TITULOSISTEMA, _
                                  "CaminhoBMP", _
                                  txtRegPathBmp.Text
  objRegister.SalvarChaveRegistro TITULOSISTEMA, _
                                  "CaminhoIcons", _
                                  txtRegPathIcons.Text
  objRegister.SalvarChaveRegistro TITULOSISTEMA, _
                                  "CaminhoReport", _
                                  txtRegPathRpt.Text
  objRegister.SalvarChaveRegistro TITULOSISTEMA, _
                                  "Empresa", _
                                  txtRegEmpresa.Text
  objRegister.SalvarChaveRegistro TITULOSISTEMA, _
                                  "BMP", _
                                  txtRegBMP.Text
  objRegister.InicializaRegister TITULOSISTEMA, _
                                 gsReportPath, _
                                 gsAppPath, _
                                 gsNomeUsu, _
                                 gsNomeEmpresa, _
                                 gsBMPPath, _
                                 gsIconsPath, _
                                 gsBMP, _
                                 gsPathBackup, _
                                 gsNomeServidorBD, _
                                 ConnectRpt
  cmdRegister.Caption = ">>"
  Me.Height = intHeightNormal
  txtRegBd.TabStop = False
  If Len(txtUsuario) = 0 Then
    SetarFoco txtUsuario
  Else
    SetarFoco txtSenha
  End If
  Set objRegister = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

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
  Me.Height = 2880
  intHeightNormal = 2880
  intHeightMax = 6315
  LerFiguras Me, tpBmp_Login, Me.cmdOk, Me.cmdCancelar
  'Carrega dados
  txtRegBd.Text = gsNomeServidorBD
  txtRegPathApp.Text = gsAppPath
  txtRegPathBackup.Text = gsPathBackup
  txtRegPathBmp.Text = gsBMPPath
  txtRegPathIcons.Text = gsIconsPath
  txtRegPathRpt.Text = gsReportPath
  txtRegEmpresa.Text = gsNomeEmpresa
  txtRegBMP.Text = gsBMP
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
