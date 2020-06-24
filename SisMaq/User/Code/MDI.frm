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
      Left            =   1020
      Top             =   4200
   End
   Begin VB.Timer tmrUnidade 
      Interval        =   60000
      Left            =   1020
      Top             =   3510
   End
   Begin VB.Timer tmrServMovCaixa 
      Left            =   1020
      Top             =   2820
   End
   Begin VB.Timer tmrServDiaria 
      Left            =   1020
      Top             =   2130
   End
   Begin VB.Timer tmrServDesp 
      Left            =   1020
      Top             =   1410
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2340
      Top             =   1410
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
            TextSave        =   "21/3/2011"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "22:59"
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
         Caption         =   "&Papel de Parede"
         Index           =   2
      End
      Begin VB.Menu mnuArquivo 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuArquivo 
         Caption         =   "&Sair"
         Index           =   4
      End
   End
   Begin VB.Menu snuGerencia 
      Caption         =   "&Gerencia"
      Index           =   0
      Begin VB.Menu mnuGerencia 
         Caption         =   "&Operacional"
         Index           =   0
         Begin VB.Menu mnuGerOper 
            Caption         =   "&Caixa/Gerencia"
            Index           =   0
         End
         Begin VB.Menu mnuGerOper 
            Caption         =   "&Atendente"
            Index           =   1
         End
         Begin VB.Menu mnuGerOper 
            Caption         =   "A&rrecadador"
            Index           =   2
         End
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "&Turno"
         Index           =   1
         Begin VB.Menu mnuGerTurno 
            Caption         =   "&Turno"
            Index           =   0
         End
         Begin VB.Menu mnuGerTurno 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuGerTurno 
            Caption         =   "&Período"
            Index           =   2
         End
         Begin VB.Menu mnuGerTurno 
            Caption         =   "&Dia da semana"
            Index           =   3
         End
         Begin VB.Menu mnuGerTurno 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuGerTurno 
            Caption         =   "&Fluxo de Caixa"
            Index           =   5
         End
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "&Equipamento"
         Index           =   2
         Begin VB.Menu mnuGerEquip 
            Caption         =   "&Tipo"
            Index           =   0
         End
         Begin VB.Menu mnuGerEquip 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuGerEquip 
            Caption         =   "&Série/Equipamento"
            Index           =   2
         End
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "&Despesa/Financeiro"
         Index           =   3
         Begin VB.Menu mnuGerDespesas 
            Caption         =   "&Grupo/Sub Grupo"
            Index           =   0
         End
         Begin VB.Menu mnuGerDespesas 
            Caption         =   "&Forma de Pagamento"
            Index           =   1
         End
         Begin VB.Menu mnuGerDespesas 
            Caption         =   "&Livro"
            Index           =   2
         End
         Begin VB.Menu mnuGerDespesas 
            Caption         =   "&Banco"
            Index           =   3
         End
         Begin VB.Menu mnuGerDespesas 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuGerDespesas 
            Caption         =   "&Despesa"
            Index           =   5
         End
         Begin VB.Menu mnuGerDespesas 
            Caption         =   "&Despesa a Pagar"
            Index           =   6
         End
         Begin VB.Menu mnuGerDespesas 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnuGerDespesas 
            Caption         =   "&Tipo de Pagamento"
            Index           =   8
         End
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "&Funcionário"
         Index           =   4
         Begin VB.Menu mnuGerFuncionario 
            Caption         =   "&Funcionário"
            Index           =   0
         End
      End
   End
   Begin VB.Menu snuCaixa 
      Caption         =   "&Caixa"
      Index           =   0
      Begin VB.Menu mnuCaixa 
         Caption         =   "&Operacional"
         Index           =   0
         Begin VB.Menu mnuCaiOper 
            Caption         =   "&Caixa"
            Index           =   0
         End
         Begin VB.Menu mnuCaiOper 
            Caption         =   "&Atendente"
            Index           =   1
         End
         Begin VB.Menu mnuCaiOper 
            Caption         =   "A&rrecadador"
            Index           =   2
         End
      End
      Begin VB.Menu mnuCaixa 
         Caption         =   "&Turno"
         Index           =   1
         Begin VB.Menu mnuCaiTurno 
            Caption         =   "&Turno"
            Index           =   0
         End
      End
      Begin VB.Menu mnuCaixa 
         Caption         =   "&Despesa/Financeiro"
         Index           =   2
         Begin VB.Menu mnuCaiDespesas 
            Caption         =   "&Despesa"
            Index           =   0
         End
         Begin VB.Menu mnuCaiDespesas 
            Caption         =   "&Despesa a Pagar"
            Index           =   1
         End
         Begin VB.Menu mnuCaiDespesas 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuCaiDespesas 
            Caption         =   "&Tipo de Pagamento"
            Index           =   3
         End
      End
   End
   Begin VB.Menu snuAtendente 
      Caption         =   "&Atendente"
      Index           =   0
      Begin VB.Menu mnuAtendente 
         Caption         =   "&Operacional"
         Index           =   0
         Begin VB.Menu mnuAteOper 
            Caption         =   "&Atendente"
            Index           =   0
         End
         Begin VB.Menu mnuAteOper 
            Caption         =   "A&rrecadador"
            Index           =   1
         End
      End
   End
   Begin VB.Menu snuArrecadador 
      Caption         =   "&Arrecadador"
      Index           =   0
      Begin VB.Menu mnuArrecadador 
         Caption         =   "&Operacional"
         Index           =   0
         Begin VB.Menu mnuArrOper 
            Caption         =   "A&rrecadador"
            Index           =   0
         End
         Begin VB.Menu mnuArrOper 
            Caption         =   "&Atendente"
            Index           =   1
         End
      End
   End
   Begin VB.Menu snuLeiturista 
      Caption         =   "&Leiturista"
      Index           =   0
      Begin VB.Menu mnuLeiturista 
         Caption         =   "&Operacional"
         Index           =   0
         Begin VB.Menu mnuLeiOper 
            Caption         =   "&Leiturista"
            Index           =   0
         End
      End
   End
   Begin VB.Menu snuFinanceiro 
      Caption         =   "&Financeiro"
      Index           =   0
   End
   Begin VB.Menu snuRelatorio 
      Caption         =   "Re&latórios"
      Index           =   0
      Begin VB.Menu mnuRelLeitura 
         Caption         =   "&Fechamento do Caixa"
         Index           =   0
      End
      Begin VB.Menu mnuRelLeitura 
         Caption         =   "&Fluxo de Caixa"
         Index           =   1
      End
      Begin VB.Menu mnuRelLeitura 
         Caption         =   "&Leituras Irregulares"
         Index           =   2
      End
      Begin VB.Menu mnuRelLeitura 
         Caption         =   "&Lucratividade por Máquina"
         Index           =   3
      End
      Begin VB.Menu mnuRelLeitura 
         Caption         =   "&Lucratividade por Série"
         Index           =   4
      End
      Begin VB.Menu mnuRelLeitura 
         Caption         =   "&Medições de Arrecadadores"
         Index           =   5
      End
      Begin VB.Menu mnuRelLeitura 
         Caption         =   "&Medições de Atendentes"
         Index           =   6
      End
      Begin VB.Menu mnuRelLeitura 
         Caption         =   "&Movimento por Máquina"
         Index           =   7
      End
      Begin VB.Menu mnuRelLeitura 
         Caption         =   "&Validação de medições"
         Index           =   8
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
  Dim objGeral    As busSisMaq.clsGeral
  Dim objForm     As Form
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
    Set objGeral = New busSisMaq.clsGeral
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
    Case gsCaixa, gsGerente, gsAdmin
      'frmUserOperCaiCons
      Set objForm = New SisMaq.frmUserOperCaiCons
      objForm.Status = tpStatus_Incluir
      objForm.lngTURNOATENDEPESQ = 0
      objForm.Show vbModal
      Set objForm = Nothing
    Case gsAtend
      'frmUserOperAteCons
      Set objForm = New SisMaq.frmUserOperAteCons
      objForm.Status = tpStatus_Incluir
      objForm.lngTURNOATENDEPESQ = 0
      objForm.Show vbModal
      Set objForm = Nothing
    Case gsArrec
      'frmUserOperArrCons
      Set objForm = New SisMaq.frmUserOperArrCons
      objForm.Status = tpStatus.tpStatus_Incluir
      objForm.lngTURNOARRECEPESQ = 0
      objForm.Show vbModal
      Set objForm = Nothing
    Case gsLeiturista
      'frmUserOperLeiCons
      Set objForm = New SisMaq.frmUserOperLeiCons
      objForm.Show vbModal
      Set objForm = Nothing
      
      
'    Case gsRecepcao
'      Set objForm = New SisMaq.frmUserLocacaoBot
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
'''  Dim objProtec As busSisMaq.clsProtec
'''  Dim objGeral As busSisMaq.clsGeral
'''  Set objProtec = New busSisMaq.clsProtec
'''  Set objGeral = New busSisMaq.clsGeral
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
  Case 2: frmUserPapel.Show vbModal
  Case 4: Unload Me
  End Select
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub



Private Sub mnuArrOper_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserOperArrCons.Show vbModal
  Case 1: frmUserOperAteCons.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuAteOper_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserOperAteCons.Show vbModal
  Case 1: frmUserOperArrCons.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuCaiDespesas_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0
    frmUserDespesaCtaLis.strTipo = "A" 'Administração
    frmUserDespesaCtaLis.strTipoCtaPagas = "N"
    frmUserDespesaCtaLis.Show vbModal
  Case 1
    frmUserDespesaCtaLis.strTipo = "A" 'Administração
    frmUserDespesaCtaLis.strTipoCtaPagas = "S"
    frmUserDespesaCtaLis.Show vbModal
  Case 3: frmUserTipoPgtoLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuCaiOper_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserOperCaiCons.Show vbModal
  Case 1: frmUserOperAteCons.Show vbModal
  Case 2: frmUserOperArrCons.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuCaiTurno_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserTurnoInc.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuGerDespesas_Click(Index As Integer)
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
  Case 8: frmUserTipoPgtoLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuGerEquip_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserTipoLis.Show vbModal
  Case 2: frmUserSerieLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuGerFuncionario_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0
    frmUserPessoaLis.IcPessoa = tpIcPessoa_Func
    frmUserPessoaLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuGerOper_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserOperCaiCons.Show vbModal
  Case 1: frmUserOperAteCons.Show vbModal
  Case 2: frmUserOperArrCons.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuGerTurno_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserTurnoInc.Show vbModal
  Case 2: frmUserPeriodoLis.Show vbModal
  Case 3: frmUserDiasDaSemanaLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuLeiOper_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserOperLeiCons.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuRelLeitura_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserRelFechaCaixa.Show vbModal
  Case 1: frmUserRelFluxoCaixa.Show vbModal
  Case 2: frmUserRelLeituraIrregular.Show vbModal
  Case 3: frmUserRelLucratividade.Show vbModal
  Case 4: frmUserRelLucratividadeSerie.Show vbModal
  Case 5: frmUserRelMedArrecadador.Show vbModal
  Case 6: frmUserRelMedAtendente.Show vbModal
  Case 7: frmUserRelMovMaquina.Show vbModal
  Case 8: frmUserRelValidaMed.Show vbModal
  
  End Select
  AmpN

End Sub

Private Sub snuSobre_Click()
  AmpS
  frmUserAbout.Show
  AmpN
End Sub
