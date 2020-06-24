VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8880
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrProtecao 
      Interval        =   60000
      Left            =   720
      Top             =   3030
   End
   Begin VB.Timer tmrUnidade 
      Interval        =   60000
      Left            =   720
      Top             =   2340
   End
   Begin VB.Timer tmrServMovCaixa 
      Left            =   720
      Top             =   1650
   End
   Begin VB.Timer tmrServDiaria 
      Left            =   720
      Top             =   960
   End
   Begin VB.Timer tmrServDesp 
      Left            =   720
      Top             =   240
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar stbPrinc 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4995
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10584
            MinWidth        =   10584
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   2
            TextSave        =   "29/4/2010"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "22:35"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu snuArquivo 
      Caption         =   "&Sistema"
      Begin VB.Menu mnuArquivo 
         Caption         =   "&Desconectar"
         Index           =   0
      End
      Begin VB.Menu mnuArquivo 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuArquivo 
         Caption         =   "&Funcionário"
         Index           =   2
      End
      Begin VB.Menu mnuArquivo 
         Caption         =   "&Papel de Parede"
         Index           =   4
      End
      Begin VB.Menu mnuArquivo 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuArquivo 
         Caption         =   "&Sair"
         Index           =   6
      End
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastro"
      Begin VB.Menu mnuConvenio 
         Caption         =   "Convênio"
         Begin VB.Menu snuConvenio 
            Caption         =   "Convênio"
            Index           =   0
         End
         Begin VB.Menu snuConvenio 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu snuConvenio 
            Caption         =   "Tipo de Convênio"
            Index           =   2
         End
         Begin VB.Menu snuConvenio 
            Caption         =   "&Parceiro"
            Index           =   3
         End
      End
      Begin VB.Menu mnuAssociado 
         Caption         =   "Associado"
         Begin VB.Menu snuAssociado 
            Caption         =   "Associado"
            Index           =   0
         End
         Begin VB.Menu snuAssociado 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu snuAssociado 
            Caption         =   "Tipo de Sócio"
            Index           =   2
         End
         Begin VB.Menu snuAssociado 
            Caption         =   "Profissão"
            Index           =   3
         End
         Begin VB.Menu snuAssociado 
            Caption         =   "Grau de Parentesco"
            Index           =   4
         End
         Begin VB.Menu snuAssociado 
            Caption         =   "Origem"
            Index           =   5
         End
         Begin VB.Menu snuAssociado 
            Caption         =   "Empresa"
            Index           =   6
         End
         Begin VB.Menu snuAssociado 
            Caption         =   "Linha"
            Index           =   7
         End
         Begin VB.Menu snuAssociado 
            Caption         =   "Estado Cívil"
            Index           =   8
         End
      End
      Begin VB.Menu mnuPlano 
         Caption         =   "&Plano"
         Begin VB.Menu snuPlano 
            Caption         =   "Plano Apler"
            Index           =   0
         End
      End
      Begin VB.Menu mnuCaptador 
         Caption         =   "&Captador"
         Begin VB.Menu snuCaptador 
            Caption         =   "&Captador"
            Index           =   0
         End
      End
      Begin VB.Menu mnuFuncionario 
         Caption         =   "&Funcionário"
         Begin VB.Menu snuFuncionario 
            Caption         =   "&Funcionário"
            Index           =   0
         End
         Begin VB.Menu snuFuncionario 
            Caption         =   "&Qualificação"
            Index           =   1
         End
      End
   End
   Begin VB.Menu snuFinanceiro 
      Caption         =   "&Financeiro"
      Index           =   0
      Begin VB.Menu snuFinanceiroC 
         Caption         =   "&Despesas"
         Index           =   0
         Begin VB.Menu mnuDespesas 
            Caption         =   "&Grupo/Sub Grupo"
            Index           =   0
         End
         Begin VB.Menu mnuDespesas 
            Caption         =   "&Forma de Pagamento"
            Index           =   1
         End
         Begin VB.Menu mnuDespesas 
            Caption         =   "&Livro"
            Index           =   2
         End
         Begin VB.Menu mnuDespesas 
            Caption         =   "&Banco"
            Index           =   3
         End
         Begin VB.Menu mnuDespesas 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuDespesas 
            Caption         =   "&Despesas"
            Index           =   5
         End
         Begin VB.Menu mnuDespesas 
            Caption         =   "&Despesas a Pagar"
            Index           =   6
         End
         Begin VB.Menu mnuDespesas 
            Caption         =   "&Saldo"
            Index           =   7
         End
      End
      Begin VB.Menu snuFinanceiroC 
         Caption         =   "&Receitas"
         Index           =   1
         Begin VB.Menu mnuReceitas 
            Caption         =   "&Tipo de empresa"
            Index           =   0
         End
         Begin VB.Menu mnuReceitas 
            Caption         =   "&Empresa"
            Index           =   1
         End
         Begin VB.Menu mnuReceitas 
            Caption         =   "&Contratos"
            Index           =   2
         End
         Begin VB.Menu mnuReceitas 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuReceitas 
            Caption         =   "&Receitas"
            Index           =   4
         End
         Begin VB.Menu mnuReceitas 
            Caption         =   "&Receitas a receber"
            Index           =   5
         End
      End
      Begin VB.Menu snuFinanceiroC 
         Caption         =   "&Movimentações"
         Index           =   2
         Begin VB.Menu mnuMovimentacoes 
            Caption         =   "&Contas"
            Index           =   0
         End
         Begin VB.Menu mnuMovimentacoes 
            Caption         =   "&Zerar Contas"
            Index           =   1
         End
         Begin VB.Menu mnuMovimentacoes 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuMovimentacoes 
            Caption         =   "&Movimentação"
            Index           =   3
         End
         Begin VB.Menu mnuMovimentacoes 
            Caption         =   "&Ajuste de Débito"
            Index           =   4
         End
         Begin VB.Menu mnuMovimentacoes 
            Caption         =   "&Ajuste de Crédito"
            Index           =   5
         End
      End
      Begin VB.Menu snuFinanceiroC 
         Caption         =   "&Cliente/Cheques"
         Index           =   3
         Visible         =   0   'False
      End
   End
   Begin VB.Menu snuRelatorio 
      Caption         =   "Re&latórios"
      Index           =   0
      Begin VB.Menu snuRelFinanc 
         Caption         =   "&Financeiros"
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Gerencial"
            Index           =   0
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Demonstrativo de Faturamento/Recebimento por turno"
            Index           =   1
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Demonstrativo de Faturamento por tipo pgto. por turno"
            Index           =   2
         End
      End
   End
   Begin VB.Menu snuSobre 
      Caption         =   "So&bre"
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnPrimeiraVez  As Boolean
Public objForm          As Form

Public Sub LerFiguras()
   '
   Me.Icon = LoadPicture(gsIconsPath & "Logo.ico")
   '
End Sub

Private Sub MDIForm_Activate()
  On Error GoTo trata
  Dim objGeral    As busApler.clsGeral
  '
  If App.PrevInstance Then
    MsgBox "Aplicativo já está rodando!", vbExclamation, TITULOSISTEMA
    End
  End If
  If blnPrimeiraVez Then
    '-----------------
    '------------ INICIO
    '----------------
    'Eugenio, para tirar a proteção, comente o código abaixo até antes de End Sub
    'Depois vá em Project/References e desmarque as referencias para Protec
    '---------------------------------------------------------------
    '----------------
    'Proteção do sistema
    '----------------
    Set objGeral = New busApler.clsGeral
    '----------------
    'Verifica Proteção do sistema
    '-------------------------
    'Valida primeira vez que entrou no sistema
    '----------------
    '----------------
    'Atualizar data Atual do sistema
    'Mata o arquivo fisicamene
    Set objGeral = Nothing
    '-----------------
    '------------ FIM
    '----------------
'''    If Now() > CDate("2003/10/01") Then
'''      TratarErroPrevisto "Acabou a validade desta cópia do sistema, contacte o suporte para adquirir uma nova versão.", "[frmMDI_Activqte]"
'''      End
'''    End If
    '
    If Len(Trim(gsBMP)) <> 0 Then
      If Dir(gsBMP) <> "" Then
        Me.Picture = LoadPicture(gsBMP)
      End If
    End If
    AmpN
    '
    Monta_Menu 1
    Select Case gsNivel
    Case gsPortaria
'      'Unidades
'      Set objForm = New Apler.frmUserLocacaoBot
'      objForm.QuemChamou = 0 'Chamada é de Portaria
'      objForm.Show vbModal
'      Set objForm = Nothing
'    Case gsRecepcao
'      Set objForm = New Apler.frmUserLocacaoBot
'      objForm.QuemChamou = 1 'Chamada é de Recepção
'      objForm.Show vbModal
'      Set objForm = Nothing
'    Case gsGerente
'      'frmUserLocacao.QuemChamou = 2 'Chamada é da Gerencia
'      'frmUserLocacao.Show vbModal
'    Case gsGerente, gsDiretor, gsAdmin
'      'frmUserLocacao.QuemChamou = 3 'Chamada é da Diretoria / Administração
'      'frmUserLocacao.Show vbModal
    End Select
    
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  End
End Sub

Private Sub MDIForm_Load()
  On Error Resume Next
  blnPrimeiraVez = True
  AmpS
  Me.Caption = TITULOSISTEMA & " - " & gsNomeEmpresa
  '
  LerFiguras
  '
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    If Trim(gsNivel) = gsRecepcao Or Trim(gsNivel) = gsPortaria Or Trim(gsNivel) = "" Then
      MsgBox "Você não tem autorização para sair do sistema. Para efetuar essa operação, vá em arquivo/Desconectar, depois vá em arquivo/Conectar e chame seu gerente/Diretor para entrar com a senha e sair do sistema.", vbExclamation, TITULOSISTEMA
      Cancel = True
    Else
      CapturaParametrosRegistro 3
      End
    End If
End Sub
'''
'''Private Sub tmrProtecao_Timer()
'''  On Error GoTo trata
'''  Dim objProtec As busApler.clsProtec
'''  Dim objGeral As busApler.clsGeral
'''  Set objProtec = New busApler.clsProtec
'''  Set objGeral = New busApler.clsGeral
'''  '----------------
'''  'Atualizar data Atual do sistema
'''  objProtec.Atualiza_Chave_Data_Atual objGeral.ObterConnectionString
'''  Set objProtec = Nothing
'''  Set objGeral = Nothing
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, _
'''             Err.Description, _
'''             Err.Source
'''End Sub

Private Sub mnuArquivo_Click(Index As Integer)
  On Error GoTo trata
  AmpS
  Select Case Index
  Case 0
    If frmMDI.mnuArquivo(0).Caption = "&Desconectar" Then
      frmMDI.mnuArquivo(0).Caption = "&Conectar"
      Monta_Menu 0
      '
      'Captura configurações do Usuário
      gsNomeUsu = ""
      gsNivel = ""
      '
      frmMDI.stbPrinc.Panels(1).Text = gsNomeUsu
      frmMDI.stbPrinc.Panels(2).Text = gsNivel
      '
      'Captura_Config
      'HabServDesp
      '
    Else
      frmUserLogin.QuemChamou = 1
      frmUserLogin.Show vbModal
      blnPrimeiraVez = True
      '
      frmMDI.stbPrinc.Panels(1).Text = gsNomeUsu
      frmMDI.stbPrinc.Panels(2).Text = gsNivel
      '
      'Captura_Config
      'HabServDesp
      '
    End If
  '
  Case 2
    frmUserPessoaLis.IcPessoa = tpIcPessoa_Func
    frmUserPessoaLis.Show vbModal
  Case 4: frmUserPapel.Show vbModal
  Case 6: Unload Me
  End Select
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub


Private Sub mnuDespesas_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserGrupoDespesaLis.Show vbModal
  Case 1: frmUserFormaPgtoLis.Show vbModal
  Case 2: frmUserLivroLis.Show vbModal
  Case 3: frmUserBancoLis.Show vbModal
  Case 5
    frmUserDespesaCtaLis.strTipo = "A" 'Administração
    frmUserDespesaCtaLis.strTipoCtaPagas = "N"
    frmUserDespesaCtaLis.Show vbModal
  Case 6
    frmUserDespesaCtaLis.strTipo = "A" 'Administração
    frmUserDespesaCtaLis.strTipoCtaPagas = "S"
    frmUserDespesaCtaLis.Show vbModal
  Case 7: frmUserSaldoLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuMovimentacoes_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserContaLis.Show vbModal
  Case 1
    frmUserZerarContaInc.Status = tpStatus_Alterar
    frmUserZerarContaInc.Show vbModal 'Zerar contas
  Case 3 'Movimentação
    frmUserMovimentacaoLis.strStatus = "M"
    frmUserMovimentacaoLis.Show vbModal
  Case 4 'Ajuste de Débito
    frmUserMovimentacaoLis.strStatus = "D"
    frmUserMovimentacaoLis.Show vbModal
  Case 5 'Ajuste de Crédito
    frmUserMovimentacaoLis.strStatus = "C"
    frmUserMovimentacaoLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuReceitas_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserTipoEmpresaLis.Show vbModal
  Case 1: frmUserEmpresaCtaLis.Show vbModal
  Case 2: frmUserContratoLis.Show vbModal
  Case 4
    frmUserReceitaLis.strTipo = "A" 'Administração
    frmUserReceitaLis.strTipoCtaRecebidas = "N"
    frmUserReceitaLis.Show vbModal
  Case 5
    frmUserReceitaLis.strTipo = "A" 'Administração
    frmUserReceitaLis.strTipoCtaRecebidas = "S"
    frmUserReceitaLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub snuAssociado_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserAssociadoLis.Show vbModal
  Case 2: frmUserTipoSocioLis.Show vbModal
  Case 3: frmUserProfissaoLis.Show vbModal
  Case 4: frmUserGrauParentescoLis.Show vbModal
  Case 5: frmUserOrigemLis.Show vbModal
  Case 6: frmUserEmpresaLis.Show vbModal
  Case 7: frmUserLinhaLis.Show vbModal
  Case 8: frmUserEstadoCivilLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub snuCaptador_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserCaptadorLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub snuConvenio_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserConvenioLis.Show vbModal
  Case 2: frmUserTipoConvenioLis.Show vbModal
  Case 3: frmUserParceiroLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub snuFuncionario_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0
    frmUserPessoaLis.IcPessoa = tpIcPessoa_Func
    frmUserPessoaLis.Show vbModal
  Case 1: frmUserQualificacaoLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub snuPlano_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserPlanoLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub snuSobre_Click()
  AmpS
  frmUserAbout.Show
  AmpN
End Sub
