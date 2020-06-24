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
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10584
            MinWidth        =   10584
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   2
            TextSave        =   "15/3/2016"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "23:38"
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
         Caption         =   "&Usuários"
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
   Begin VB.Menu snuGerencia 
      Caption         =   "&Gerencia"
      Index           =   0
      Begin VB.Menu mnuGerencia 
         Caption         =   "&Banco"
         Index           =   0
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "&Motorista"
         Index           =   1
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "M&arca"
         Index           =   3
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "M&odelo"
         Index           =   4
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "&Veículo"
         Index           =   5
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "A&gência"
         Index           =   7
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "&Origem/Destino"
         Index           =   8
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "&Serviço"
         Index           =   9
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "&Pacote"
         Index           =   11
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
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Demonstrativo de Recebimento por Tipo de Unidade"
            Index           =   3
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Demonstrativo de Recebimento por Unidade"
            Index           =   4
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Demonstrativo de Entrada/Recebimento por Unidade"
            Index           =   5
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Resumo Diário"
            Index           =   6
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Resumo Diário (resumido)"
            Index           =   7
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Resumo Diário de Receitas e Despesas"
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Resumo Diário de Receitas e Despesas"
            Index           =   9
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Mapa Financeiro Diário"
            Index           =   10
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Resumo de adiantamentos retirados"
            Index           =   11
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Demonstrativo de Despesas por Turno"
            Index           =   12
         End
      End
      Begin VB.Menu snuRelCartChq 
         Caption         =   "&Cartões/Cheques/Penhores"
         Begin VB.Menu mnuRelCartChq 
            Caption         =   "D&emonstrativo de Recebimento por cartão de crédito"
            Index           =   0
         End
         Begin VB.Menu mnuRelCartChq 
            Caption         =   "D&emonstrativo de Recebimento por cartão de débito"
            Index           =   1
         End
         Begin VB.Menu mnuRelCartChq 
            Caption         =   "&Recebimento por lote"
            Index           =   2
         End
         Begin VB.Menu mnuRelCartChq 
            Caption         =   "Demonstrativo de rec. de cheques  por data de receb."
            Index           =   3
         End
         Begin VB.Menu mnuRelCartChq 
            Caption         =   "Demonstrativo de rec. de penhores"
            Index           =   4
         End
      End
      Begin VB.Menu snuRelVdaCanc 
         Caption         =   "&Vendas/Cancelamentos"
         Begin VB.Menu mnuRelVdaCanc 
            Caption         =   "Cancelamento de Pedidos"
            Index           =   0
         End
         Begin VB.Menu mnuRelVdaCanc 
            Caption         =   "Cancelamento de Locações"
            Index           =   1
         End
         Begin VB.Menu mnuRelVdaCanc 
            Caption         =   "Demonstrativos de vendas Diversas &cobradas"
            Index           =   2
         End
         Begin VB.Menu mnuRelVdaCanc 
            Caption         =   "Demonstrativos de vendas &Diversas não cobradas"
            Index           =   3
         End
         Begin VB.Menu mnuRelVdaCanc 
            Caption         =   "Demonstrativos de vendas &Extras fora da unidade"
            Index           =   4
         End
         Begin VB.Menu mnuRelVdaCanc 
            Caption         =   "Demonstrativos de vendas &Extras na unidade"
            Index           =   5
         End
      End
      Begin VB.Menu snuRelCont 
         Caption         =   "&Controles"
         Begin VB.Menu mnuRelCont 
            Caption         =   "Demonstrativo de Gorjetas por Turno"
            Index           =   0
         End
         Begin VB.Menu mnuRelCont 
            Caption         =   "Demonstrativo de trabalho das camareiras"
            Index           =   1
         End
         Begin VB.Menu mnuRelCont 
            Caption         =   "&Frequência"
            Index           =   2
         End
         Begin VB.Menu mnuRelCont 
            Caption         =   "Relação de clientes por turno/placa/cpf"
            Index           =   3
         End
         Begin VB.Menu mnuRelCont 
            Caption         =   "Demonstrativo de Ocupação"
            Index           =   4
         End
         Begin VB.Menu mnuRelCont 
            Caption         =   "Demonstrativo de &Funcionários"
            Index           =   5
         End
      End
      Begin VB.Menu snuRelPromo 
         Caption         =   "Promoções"
         Begin VB.Menu mnuRelPromo 
            Caption         =   "Demonstrativos de &Desconto/Promoções/Cortesias"
            Index           =   0
         End
         Begin VB.Menu mnuRelPromo 
            Caption         =   "Demonstrativo de &promoções"
            Index           =   1
         End
      End
      Begin VB.Menu snuRelDemoCon 
         Caption         =   "&Demonstrativo de Consumo"
         Begin VB.Menu mnuRelDemoCon 
            Caption         =   "&Campeões de Vendas - Tipo / Grupo"
            Index           =   0
         End
         Begin VB.Menu mnuRelDemoCon 
            Caption         =   "&Campeões de Vendas - Tipo"
            Index           =   1
         End
         Begin VB.Menu mnuRelDemoCon 
            Caption         =   "Movimento de Copa/Cozinha/Frigobar/Outros"
            Index           =   2
         End
         Begin VB.Menu mnuRelDemoCon 
            Caption         =   "Relatório Sintético de Estoque"
            Index           =   3
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
  Dim objGeral    As busElite.clsGeral
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
    Set objGeral = New busElite.clsGeral
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
'      Set objForm = New Elite.frmUserLocacaoBot
'      objForm.QuemChamou = 0 'Chamada é de Portaria
'      objForm.Show vbModal
'      Set objForm = Nothing
'    Case gsRecepcao
'      Set objForm = New Elite.frmUserLocacaoBot
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
    frmGerencial.Show vbModal
    
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
'''  Dim objProtec As busElite.clsProtec
'''  Dim objGeral As busElite.clsGeral
'''  Set objProtec = New busElite.clsProtec
'''  Set objGeral = New busElite.clsGeral
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



Private Sub mnuGerencia_Click(Index As Integer)
  On Error GoTo trata
  AmpS
  Select Case Index
  Case 0: frmBancoLis.Show vbModal
  Case 1
    frmUserPessoaLis.IcPessoa = tpIcPessoa_Mot
    frmUserPessoaLis.Show vbModal
  Case 3: frmMarcaLis.Show vbModal
  Case 4: frmModeloLis.Show vbModal
  Case 5: frmVeiculoLis.Show vbModal
  Case 7: frmAgenciaLis.Show vbModal
  Case 8: frmOrigemLis.Show vbModal
  Case 9: frmServicoLis.Show vbModal
  Case 11: frmGerencial.Show vbModal
  End Select
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub snuSobre_Click()
  AmpS
  frmUserAbout.Show
  AmpN
End Sub
