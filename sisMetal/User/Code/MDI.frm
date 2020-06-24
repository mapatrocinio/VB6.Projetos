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
      Left            =   1710
      Top             =   3660
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
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10584
            MinWidth        =   10584
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   2
            TextSave        =   "6/7/2015"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "18:45"
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
         Caption         =   "&Papel de Parede"
         Index           =   2
      End
      Begin VB.Menu mnuArquivo 
         Caption         =   "&Configurações"
         Index           =   3
      End
      Begin VB.Menu mnuArquivo 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuArquivo 
         Caption         =   "&Sair"
         Index           =   5
      End
   End
   Begin VB.Menu snuCadastro 
      Caption         =   "&Cadastro"
      Index           =   0
      Begin VB.Menu snuLoja 
         Caption         =   "&Loja"
         Index           =   0
         Begin VB.Menu mnuLoja 
            Caption         =   "&Fábrica"
            Index           =   0
         End
         Begin VB.Menu mnuLoja 
            Caption         =   "F&ilial"
            Index           =   1
         End
         Begin VB.Menu mnuLoja 
            Caption         =   "&Anodizadora"
            Index           =   2
         End
         Begin VB.Menu mnuLoja 
            Caption         =   "F&ornecedor"
            Index           =   3
         End
      End
      Begin VB.Menu snuInsumo 
         Caption         =   "&Insumo"
         Index           =   0
         Begin VB.Menu mnuInsumo 
            Caption         =   "&Cor"
            Index           =   0
         End
         Begin VB.Menu mnuInsumo 
            Caption         =   "&Linha"
            Index           =   1
         End
         Begin VB.Menu mnuInsumo 
            Caption         =   "L&inha-Perfil"
            Index           =   2
         End
         Begin VB.Menu mnuInsumo 
            Caption         =   "&Metragem"
            Index           =   3
         End
         Begin VB.Menu mnuInsumo 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuInsumo 
            Caption         =   "&Perfil"
            Index           =   5
         End
         Begin VB.Menu mnuInsumo 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuInsumo 
            Caption         =   "&Grupo"
            Index           =   7
         End
         Begin VB.Menu mnuInsumo 
            Caption         =   "&Unidade de medida"
            Index           =   8
         End
         Begin VB.Menu mnuInsumo 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu mnuInsumo 
            Caption         =   "&Acessório"
            Index           =   10
         End
         Begin VB.Menu mnuInsumo 
            Caption         =   "-"
            Index           =   11
         End
         Begin VB.Menu mnuInsumo 
            Caption         =   "&IPI"
            Index           =   12
         End
         Begin VB.Menu mnuInsumo 
            Caption         =   "I&CMS"
            Index           =   13
         End
         Begin VB.Menu mnuInsumo 
            Caption         =   "&Família de Produto"
            Index           =   14
         End
         Begin VB.Menu mnuInsumo 
            Caption         =   "&Referência Produto"
            Index           =   15
         End
         Begin VB.Menu mnuInsumo 
            Caption         =   "Grupo &Produto"
            Index           =   16
         End
         Begin VB.Menu mnuInsumo 
            Caption         =   "-"
            Index           =   17
         End
         Begin VB.Menu mnuInsumo 
            Caption         =   "&Produto"
            Index           =   18
         End
      End
      Begin VB.Menu mnuCadastro 
         Caption         =   "&Usuários"
         Index           =   0
      End
   End
   Begin VB.Menu snuCompra 
      Caption         =   "&Compra"
      Index           =   0
      Begin VB.Menu mnuGerencial 
         Caption         =   "&Gerencial"
         Index           =   0
      End
   End
   Begin VB.Menu snuVenda 
      Caption         =   "&Venda"
      Index           =   0
      Begin VB.Menu mnuVenda 
         Caption         =   "&Tipo de documento"
         Index           =   0
      End
      Begin VB.Menu mnuVenda 
         Caption         =   "&Cliente"
         Index           =   1
      End
      Begin VB.Menu mnuVenda 
         Caption         =   "&Empresa"
         Index           =   2
      End
      Begin VB.Menu mnuVenda 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuVenda 
         Caption         =   "Tipo de &pagamento"
         Index           =   4
      End
      Begin VB.Menu mnuVenda 
         Caption         =   "Tipo de &estorno"
         Index           =   5
      End
      Begin VB.Menu mnuVenda 
         Caption         =   "Tipo de &venda"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVenda 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuVenda 
         Caption         =   "&Pedido"
         Index           =   8
      End
      Begin VB.Menu mnuVenda 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuVenda 
         Caption         =   "&Cartão de débito"
         Index           =   10
      End
      Begin VB.Menu mnuVenda 
         Caption         =   "&Cartão de crédito"
         Index           =   11
      End
      Begin VB.Menu mnuVenda 
         Caption         =   "&Banco"
         Index           =   12
      End
      Begin VB.Menu mnuVenda 
         Caption         =   "&Caixa"
         Index           =   13
      End
   End
   Begin VB.Menu snuEstoquePrinc 
      Caption         =   "&Estoque"
      Index           =   0
      Begin VB.Menu mnuAjuste 
         Caption         =   "&Tipo de Ajuste"
         Index           =   0
      End
      Begin VB.Menu mnuAjuste 
         Caption         =   "&Ajuste de estoque"
         Index           =   1
      End
      Begin VB.Menu mnuAjuste 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuAjuste 
         Caption         =   "&Documento de entrada"
         Index           =   3
      End
      Begin VB.Menu mnuAjuste 
         Caption         =   "&Entrada de material"
         Index           =   4
      End
      Begin VB.Menu mnuAjuste 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuAjuste 
         Caption         =   "Documento de &saída"
         Index           =   6
      End
      Begin VB.Menu mnuAjuste 
         Caption         =   "&Saída/Transf. de material"
         Index           =   7
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
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuDespesas 
            Caption         =   "&Despesas/Receitas"
            Index           =   4
         End
         Begin VB.Menu mnuDespesas 
            Caption         =   "&Contas a Pagar"
            Index           =   5
         End
      End
      Begin VB.Menu snuFinanceiroC 
         Caption         =   "&Relatórios"
         Index           =   1
         Begin VB.Menu snuRelFinancCta 
            Caption         =   "&Financeiros"
            Begin VB.Menu mnuRelFinancCta 
               Caption         =   "&Demonstrativo de Contas Pagas/à pagar"
               Index           =   0
            End
            Begin VB.Menu mnuRelFinancCta 
               Caption         =   "Demonstrativo &Resumo Geral de Despesas"
               Index           =   1
            End
            Begin VB.Menu mnuRelFinancCta 
               Caption         =   "&Listagem de Grupo / Sub Grupo"
               Index           =   2
            End
         End
      End
      Begin VB.Menu snuFinanceiroC 
         Caption         =   "&Cliente/Cheques"
         Index           =   2
      End
   End
   Begin VB.Menu snuRelatorio 
      Caption         =   "Re&latório"
      Index           =   0
      Begin VB.Menu snuRelEstoque 
         Caption         =   "&Estoque"
         Begin VB.Menu mnuRelEstoque 
            Caption         =   "&Estoque de Perfis"
            Index           =   0
         End
         Begin VB.Menu mnuRelEstoque 
            Caption         =   "&Estoque de Acessórios"
            Index           =   1
         End
         Begin VB.Menu mnuRelEstoque 
            Caption         =   "&Saldo de perfis na anodização"
            Index           =   2
         End
         Begin VB.Menu mnuRelEstoque 
            Caption         =   "Saldo de perfis no &fornecedor"
            Index           =   3
         End
         Begin VB.Menu mnuRelEstoque 
            Caption         =   "&Saldo na anodizadora"
            Index           =   4
         End
         Begin VB.Menu mnuRelEstoque 
            Caption         =   "Estoque de &Produtos"
            Index           =   5
         End
         Begin VB.Menu mnuRelEstoque 
            Caption         =   "Relatório de &Vendas"
            Index           =   6
         End
         Begin VB.Menu mnuRelEstoque 
            Caption         =   "&Demonstrativo de Faturamento"
            Index           =   7
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
  Dim objGeral    As busSisMetal.clsGeral
  Dim objProtec   As busSisMetal.clsProtec
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
    Set objProtec = New busSisMetal.clsProtec
    Set objGeral = New busSisMetal.clsGeral
    '----------------
    'Verifica Proteção do sistema
    '-------------------------
    'Valida primeira vez que entrou no sistema
    If Not objProtec.Valida_Primeira_Vez(objGeral.ObterConnectionString, App.Path) Then
      Set objProtec = Nothing
      Set objGeral = Nothing
      End
      Exit Sub
    End If
    'Válida Equipamento
    If Not objProtec.Valida_Estacao(objGeral.ObterConnectionString) Then
      Set objProtec = Nothing
      Set objGeral = Nothing
      End
      Exit Sub
    End If
    '----------------
    'Valida se sistema expirou
    If Not objProtec.Valida_Chave(objGeral.ObterConnectionString, "S", gsNivel) Then
      Set objGeral = Nothing
      Set objProtec = Nothing
      End
      Exit Sub
    End If
    '----------------
    'Atualizar data Atual do sistema
    objProtec.Atualiza_Chave_Data_Atual objGeral.ObterConnectionString
    'Mata o arquivo fisicamene
    objProtec.Trata_Arquivo_Fisico App.Path
    Set objProtec = Nothing
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
    Case gsLoja
      'Gerencial Pedido
      Set objForm = New SisMetal.frmGerencialPed
      objForm.Show vbModal
      Set objForm = Nothing
    Case gsCaixa
      'Gerencial Recebimento
      Set objForm = New SisMetal.frmGerencialRec
      objForm.Show vbModal
      Set objForm = Nothing
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
    'If Trim(gsNivel) = gsRecepcao Or Trim(gsNivel) = gsPortaria Or Trim(gsNivel) = "" Then
    '  MsgBox "Você não tem autorização para sair do sistema. Para efetuar essa operação, vá em arquivo/Desconectar, depois vá em arquivo/Conectar e chame seu gerente/Diretor para entrar com a senha e sair do sistema.", vbExclamation, TITULOSISTEMA
    '  Cancel = True
    'Else
      CapturaParametrosRegistro 3
      End
    'End If
End Sub
'''
'''Private Sub tmrProtecao_Timer()
'''  On Error GoTo trata
'''  Dim objProtec As busSisMetal.clsProtec
'''  Dim objGeral As busSisMetal.clsGeral
'''  Set objProtec = New busSisMetal.clsProtec
'''  Set objGeral = New busSisMetal.clsGeral
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

Private Sub mnuAjuste_Click(Index As Integer)
  On Error GoTo trata
  AmpS
  Select Case Index
  Case 0: frmTipoAjusteLis.Show vbModal
  Case 1: frmAjusteLis.Show vbModal
  Case 3: frmDocumentoEntradaLis.Show vbModal
  Case 4: frmEntradaMaterialLis.Show vbModal
  Case 6: frmDocumentoSaidaLis.Show vbModal
  Case 7: frmSaidaMaterialLis.Show vbModal
  End Select
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

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
      Captura_Config
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
      Captura_Config
      'HabServDesp
      '
    End If
  '
  Case 2: frmUserPapel.Show vbModal
  Case 3: frmConfiguracao.Show vbModal
  
  Case 5: Unload Me
  End Select
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub



Private Sub mnuCadastro_Click(Index As Integer)
  On Error GoTo trata
  AmpS
  Select Case Index
  Case 0
    frmUserPessoaLis.IcPessoa = tpIcPessoa_Func
    frmUserPessoaLis.Show vbModal
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
  Case 0: frmGrupoDespesaLis.Show vbModal
  Case 1: frmFormaPgtoLis.Show vbModal
  Case 2: frmLivroLis.Show vbModal
  Case 4
    frmDespesaCtaLis.strTipo = "A" 'Administração
    frmDespesaCtaLis.strTipoCtaPagas = "N"
    frmDespesaCtaLis.Show vbModal
  Case 5
    frmDespesaCtaLis.strTipo = "A" 'Administração
    frmDespesaCtaLis.strTipoCtaPagas = "S"
    frmDespesaCtaLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuGerencial_Click(Index As Integer)
  On Error GoTo trata
  AmpS
  Select Case Index
  Case 0: frmGerencial.Show vbModal
  End Select
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub mnuInsumo_Click(Index As Integer)
  On Error GoTo trata
  Dim objFormInsumo As SisMetal.frmInsumoLis
  AmpS
  Select Case Index
  Case 0: frmCorLis.Show vbModal
  Case 1: frmLinhaLis.Show vbModal
  Case 2: frmLinhaPerfilLis.Show vbModal
  Case 3: frmVaraLis.Show vbModal
  Case 5
    Set objFormInsumo = New SisMetal.frmInsumoLis
    objFormInsumo.intTipoInsumo = tpInsumo.tpInsumo_Perfil
    objFormInsumo.Show vbModal
    Set objFormInsumo = Nothing
  Case 7: frmGrupoLis.Show vbModal
  Case 8: frmEmbalagemLis.Show vbModal
  Case 10
    Set objFormInsumo = New SisMetal.frmInsumoLis
    objFormInsumo.intTipoInsumo = tpInsumo.tpInsumo_Acessorio
    objFormInsumo.Show vbModal
    Set objFormInsumo = Nothing
    
  Case 12: frmIPILis.Show vbModal
  Case 13: frmICMSLis.Show vbModal
  Case 14: frmFamiliaProdutoLis.Show vbModal
  Case 15: frmReferenciaProdutoLis.Show vbModal
    
  Case 16: frmGrupoProdutoLis.Show vbModal
  Case 18
    Set objFormInsumo = New SisMetal.frmInsumoLis
    objFormInsumo.intTipoInsumo = tpInsumo.tpInsumo_Produto
    objFormInsumo.Show vbModal
    Set objFormInsumo = Nothing
  End Select
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub mnuLoja_Click(Index As Integer)
  On Error GoTo trata
  Dim objFormLoja As SisMetal.frmLojaLis
  Set objFormLoja = New SisMetal.frmLojaLis
  AmpS
  Select Case Index
  Case 0: objFormLoja.intTipoLoja = tpLoja.tpLoja_Fabrica
  Case 1: objFormLoja.intTipoLoja = tpLoja.tpLoja_Filial
  Case 2: objFormLoja.intTipoLoja = tpLoja.tpLoja_Anodizadora
  Case 3: objFormLoja.intTipoLoja = tpLoja.tpLoja_Fornecedor
  End Select
  objFormLoja.Show vbModal
  AmpN
  Set objFormLoja = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub


Private Sub mnuRelEstoque_Click(Index As Integer)
  On Error GoTo trata
  AmpS
  Select Case Index
  Case 0: frmRelSaldoPerfil.Show vbModal
  Case 1: frmRelSaldoAcessorio.Show vbModal
  Case 2: frmRelSaldoAnodGeral.Show vbModal
  Case 3: frmRelSaldoForn.Show vbModal
  Case 4: frmRelSaldoAnod.Show vbModal
  Case 5: frmRelEstoqueProd.Show vbModal
  Case 6: frmRelItensVendas.Show vbModal
  Case 7: frmRelDemoFaturamento.Show vbModal

  End Select
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub mnuRelFinancCta_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmRelContas.Show vbModal
  Case 1: frmRelResumoDespesas.Show vbModal
  Case 2: frmRelGrupoSubGrupo.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuVenda_Click(Index As Integer)
  On Error GoTo trata
  Dim objFormLoja As SisMetal.frmLojaLis
  Dim objGerRec As SisMetal.frmGerencialRec
  AmpS
  Select Case Index
  Case 0: frmTipoDocumentoLis.Show vbModal
  Case 1: frmFichaClienteLis.Show vbModal
  Case 2
    Set objFormLoja = New SisMetal.frmLojaLis
    objFormLoja.intTipoLoja = tpLoja.tpLoja_Empresa
    objFormLoja.Show vbModal
    Set objFormLoja = Nothing
  Case 4: frmTipoPagamentoLis.Show vbModal
  Case 5: frmTipoEstornoLis.Show vbModal
  Case 6: frmTipoVendaLis.Show vbModal
  Case 8: frmGerencialPed.Show vbModal
  
  Case 10: frmCartaoDebitoLis.Show vbModal
  Case 11: frmCartaoLis.Show vbModal
  Case 12: frmBancoLis.Show vbModal
  Case 13
    Set objGerRec = New SisMetal.frmGerencialRec
    objGerRec.Show vbModal
    Set objGerRec = Nothing
  
  
  End Select
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub


Private Sub snuFinanceiroC_Click(Index As Integer)
  Select Case Index
'''  Case 3: frmUserClienteLis.Show
  End Select
End Sub

Private Sub snuSobre_Click()
  AmpS
  frmUserAbout.Show
  AmpN
End Sub
