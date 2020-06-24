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
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10584
            MinWidth        =   10584
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   2
            TextSave        =   "16/9/2012"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "06:43"
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
   Begin VB.Menu snuDiretoria 
      Caption         =   "&Diretoria"
      Index           =   0
      Begin VB.Menu mnuDiretoria 
         Caption         =   "&GR"
         Index           =   0
         Begin VB.Menu mnuDirGR 
            Caption         =   "&GR"
            Index           =   0
         End
         Begin VB.Menu mnuDirGR 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuDirGR 
            Caption         =   "&Paciente"
            Index           =   2
         End
         Begin VB.Menu mnuDirGR 
            Caption         =   "&Prestador"
            Index           =   3
         End
         Begin VB.Menu mnuDirGR 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuDirGR 
            Caption         =   "&Procedimento"
            Index           =   5
         End
         Begin VB.Menu mnuDirGR 
            Caption         =   "&Tipo de Procedimento"
            Index           =   6
         End
         Begin VB.Menu mnuDirGR 
            Caption         =   "&Especialidade"
            Index           =   7
         End
      End
      Begin VB.Menu mnuDiretoria 
         Caption         =   "&Turno"
         Index           =   1
         Begin VB.Menu mnuDirTurno 
            Caption         =   "&Turno"
            Index           =   0
         End
         Begin VB.Menu mnuDirTurno 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuDirTurno 
            Caption         =   "&Período"
            Index           =   2
         End
         Begin VB.Menu mnuDirTurno 
            Caption         =   "&Dia da semana"
            Index           =   3
         End
      End
      Begin VB.Menu mnuDiretoria 
         Caption         =   "&Sala"
         Index           =   2
         Begin VB.Menu mnuDirSala 
            Caption         =   "&Sala"
            Index           =   0
         End
         Begin VB.Menu mnuDirSala 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuDirSala 
            Caption         =   "&Prédio"
            Index           =   2
         End
      End
      Begin VB.Menu mnuDiretoria 
         Caption         =   "&Financeiro"
         Index           =   3
         Begin VB.Menu mnuDirFinanc 
            Caption         =   "&Banco"
            Index           =   0
         End
         Begin VB.Menu mnuDirFinanc 
            Caption         =   "&Cartão"
            Index           =   1
         End
         Begin VB.Menu mnuDirFinanc 
            Caption         =   "Cartão de &Débito/Convênio"
            Index           =   2
         End
      End
      Begin VB.Menu mnuDiretoria 
         Caption         =   "&Despesa"
         Index           =   4
         Begin VB.Menu mnuDirDespesa 
            Caption         =   "&Grupo/Sub Grupo"
            Index           =   0
         End
         Begin VB.Menu mnuDirDespesa 
            Caption         =   "&Forma de Pagamento"
            Index           =   1
         End
         Begin VB.Menu mnuDirDespesa 
            Caption         =   "&Livro"
            Index           =   2
         End
         Begin VB.Menu mnuDirDespesa 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuDirDespesa 
            Caption         =   "&Despesas/Receitas"
            Index           =   4
         End
      End
      Begin VB.Menu mnuDiretoria 
         Caption         =   "&Configuração"
         Index           =   5
         Begin VB.Menu mnuDirConfig 
            Caption         =   "&Registrar chave"
            Index           =   0
         End
         Begin VB.Menu mnuDirConfig 
            Caption         =   "&Configuração"
            Index           =   1
         End
      End
   End
   Begin VB.Menu snuGerencia 
      Caption         =   "&Gerencia"
      Index           =   0
      Begin VB.Menu mnuGerencia 
         Caption         =   "&GR"
         Index           =   0
         Begin VB.Menu mnuGerGR 
            Caption         =   "&GR"
            Index           =   0
         End
         Begin VB.Menu mnuGerGR 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuGerGR 
            Caption         =   "&Paciente"
            Index           =   2
         End
         Begin VB.Menu mnuGerGR 
            Caption         =   "&Prestador"
            Index           =   3
         End
         Begin VB.Menu mnuGerGR 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuGerGR 
            Caption         =   "&Procedimento"
            Index           =   5
         End
         Begin VB.Menu mnuGerGR 
            Caption         =   "&Tipo de Procedimento"
            Index           =   6
         End
         Begin VB.Menu mnuGerGR 
            Caption         =   "&Especialidade"
            Index           =   7
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
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "&Sala"
         Index           =   2
         Begin VB.Menu mnuGerSala 
            Caption         =   "&Sala"
            Index           =   0
         End
         Begin VB.Menu mnuGerSala 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuGerSala 
            Caption         =   "&Prédio"
            Index           =   2
         End
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "&Financeiro"
         Index           =   3
         Begin VB.Menu mnuGerFinanc 
            Caption         =   "&Banco"
            Index           =   0
         End
         Begin VB.Menu mnuGerFinanc 
            Caption         =   "&Cartão"
            Index           =   1
         End
         Begin VB.Menu mnuGerFinanc 
            Caption         =   "Cartão de &Débito/Convênio"
            Index           =   2
         End
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "&Despesa"
         Index           =   4
         Begin VB.Menu mnuGerDespesa 
            Caption         =   "&Grupo/Sub Grupo"
            Index           =   0
         End
         Begin VB.Menu mnuGerDespesa 
            Caption         =   "&Forma de Pagamento"
            Index           =   1
         End
         Begin VB.Menu mnuGerDespesa 
            Caption         =   "&Livro"
            Index           =   2
         End
         Begin VB.Menu mnuGerDespesa 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuGerDespesa 
            Caption         =   "&Despesas/Receitas"
            Index           =   4
         End
      End
   End
   Begin VB.Menu snuCaixa 
      Caption         =   "&Caixa"
      Index           =   0
      Begin VB.Menu mnuCaixa 
         Caption         =   "&GR"
         Index           =   0
      End
      Begin VB.Menu mnuCaixa 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuCaixa 
         Caption         =   "&Paciente"
         Index           =   2
      End
      Begin VB.Menu mnuCaixa 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuCaixa 
         Caption         =   "&Turno"
         Index           =   4
      End
   End
   Begin VB.Menu snuLaboratorio 
      Caption         =   "&Laboratório"
      Index           =   0
      Begin VB.Menu mnuLaboratorio 
         Caption         =   "&GR"
         Index           =   0
      End
      Begin VB.Menu mnuLaboratorio 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuLaboratorio 
         Caption         =   "&Paciente"
         Index           =   2
      End
      Begin VB.Menu mnuLaboratorio 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuLaboratorio 
         Caption         =   "&Turno"
         Index           =   4
      End
   End
   Begin VB.Menu snuFinanceiro 
      Caption         =   "&Financeiro"
      Index           =   0
      Begin VB.Menu mnuFinanceiro 
         Caption         =   "GR"
         Index           =   0
      End
      Begin VB.Menu mnuFinanceiro 
         Caption         =   "&Pagar Prestadores"
         Index           =   1
         Begin VB.Menu mnuPagarPrest 
            Caption         =   "GR &Paga"
            Index           =   0
         End
         Begin VB.Menu mnuPagarPrest 
            Caption         =   "GR &Paga a tec. RX"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPagarPrest 
            Caption         =   "GR &Paga a dono RX"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPagarPrest 
            Caption         =   "GR &Paga a dono Ultrason"
            Index           =   3
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuFinanceiro 
         Caption         =   "&Cancelamento de GR"
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
         Begin VB.Menu mnuCancFinan 
            Caption         =   "&Pontual"
            Index           =   0
         End
         Begin VB.Menu mnuCancFinan 
            Caption         =   "&Automatico"
            Index           =   1
         End
      End
      Begin VB.Menu mnuFinanceiro 
         Caption         =   "Con&sulta"
         Index           =   3
         Begin VB.Menu mnuConcFinan 
            Caption         =   "&GR"
            Index           =   0
         End
      End
   End
   Begin VB.Menu snuArquivista 
      Caption         =   "&Arquivista"
      Index           =   0
      Begin VB.Menu mnuArquivista 
         Caption         =   "&Arquivo"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuArquivista 
         Caption         =   "&GR"
         Index           =   1
      End
      Begin VB.Menu mnuArquivista 
         Caption         =   "&Paciente"
         Index           =   2
      End
   End
   Begin VB.Menu snuAtendimento 
      Caption         =   "&Atendimento"
      Index           =   0
      Begin VB.Menu mnuAtendimento 
         Caption         =   "&GR"
         Index           =   0
      End
   End
   Begin VB.Menu snuRelatorio 
      Caption         =   "Re&latórios"
      Index           =   0
      Begin VB.Menu snuRelFinanc 
         Caption         =   "&Financeiros"
         Index           =   0
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Pagamento realizado a prestadores"
            Index           =   0
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Pagamento realizado a prestadores (din/cart)"
            Index           =   1
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Pagamento realizado a prestadores por cartão"
            Index           =   2
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Pagamento realizado a prestadores por cartão consolidado"
            Index           =   3
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Prestador x Procedimento"
            Index           =   4
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Prestador x Especialidade"
            Index           =   5
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Atendimento Caixa"
            Index           =   6
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Demonstrativo de Contas Pagas / A pagar"
            Index           =   7
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "Demonstrativo &Resumo Geral de Receitas"
            Index           =   8
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Gr´s por Atendente"
            Index           =   9
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Gr´s não paga a prestadores"
            Index           =   10
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Gr´s canceladas (detalhado)"
            Index           =   11
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Gr´s canceladas (resumido)"
            Index           =   12
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Novos Pacientes"
            Index           =   13
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Novos Pacientes (consolidado)"
            Index           =   14
         End
      End
      Begin VB.Menu snuRelFinanc 
         Caption         =   "&Gerenciais"
         Index           =   1
         Begin VB.Menu mnuRelGerenc 
            Caption         =   "&Gerencial"
            Index           =   0
         End
         Begin VB.Menu mnuRelGerenc 
            Caption         =   "&Diário de receitas"
            Index           =   1
         End
         Begin VB.Menu mnuRelGerenc 
            Caption         =   "&Balanço"
            Index           =   2
         End
         Begin VB.Menu mnuRelGerenc 
            Caption         =   "&Pagamento realizado a prestadores por procedimento"
            Index           =   3
         End
         Begin VB.Menu mnuRelGerenc 
            Caption         =   "&Receitas"
            Index           =   4
         End
         Begin VB.Menu mnuRelGerenc 
            Caption         =   "&Receitas por Especialidade"
            Index           =   5
         End
         Begin VB.Menu mnuRelGerenc 
            Caption         =   "&Pagamento por paciente (IR)"
            Index           =   6
         End
         Begin VB.Menu mnuRelGerenc 
            Caption         =   "&Comparativo Mensal de GR´s"
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
  Dim objProtec   As busSisMed.clsProtec
  Dim objGeral    As busSisMed.clsGeral
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
    Set objProtec = New busSisMed.clsProtec
    Set objGeral = New busSisMed.clsGeral
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
    Case gsCaixa
      'GR
      frmUserGRCons.Show vbModal
    Case gsLaboratorio
      'GR
      frmUserGRCons.Show vbModal
    Case gsPrestador, gsArquivista
      'GR
      frmUserGRPreCons.Show vbModal
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
'''    If Trim(gsNivel) = gsRecepcao Or Trim(gsNivel) = gsPortaria Or Trim(gsNivel) = "" Then
'''      MsgBox "Você não tem autorização para sair do sistema. Para efetuar essa operação, vá em arquivo/Desconectar, depois vá em arquivo/Conectar e chame seu gerente/Diretor para entrar com a senha e sair do sistema.", vbExclamation, TITULOSISTEMA
'''      Cancel = True
'''    Else
      CapturaParametrosRegistro 3
      End
'''    End If
End Sub
'''
'''Private Sub tmrProtecao_Timer()
'''  On Error GoTo trata
'''  Dim objProtec As busSisMed.clsProtec
'''  Dim objGeral As busSisMed.clsGeral
'''  Set objProtec = New busSisMed.clsProtec
'''  Set objGeral = New busSisMed.clsGeral
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

Private Sub mnuArquivista_Click(Index As Integer)
  Dim objUserGRArqCons As SisMed.frmUserGRArqCons
  Dim objUserGRPreCons As SisMed.frmUserGRPreCons
  On Error GoTo trata
  AmpS
  Set objUserGRArqCons = New SisMed.frmUserGRArqCons
  Select Case Index
  Case 0
    'GR PAGA A PRESTADOR
    'objUserGRArqCons.icTipoGR = tpIcTipoGR_Prest
    'objUserGRArqCons.strGR = "GR Paga a prestador"
    objUserGRArqCons.Show vbModal
    Set objUserGRPreCons = Nothing
  Case 1
    Set objUserGRPreCons = New SisMed.frmUserGRPreCons
    objUserGRPreCons.Show vbModal
  Case 2
    frmUserProntuarioLis.IcProntuario = tpIcProntuario_Pac
    frmUserProntuarioLis.Show vbModal
  End Select
  Set objUserGRArqCons = Nothing
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
  Case 2
    frmUserProntuarioLis.IcProntuario = tpIcProntuario_Func
    frmUserProntuarioLis.Show vbModal
  Case 3: frmUserPapel.Show vbModal
  Case 5: Unload Me
  End Select
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub







Private Sub mnuAtendimento_Click(Index As Integer)
  Dim objUserGRPreCons As SisMed.frmUserGRPreCons
  AmpS
  Set objUserGRPreCons = New SisMed.frmUserGRPreCons
  Select Case Index
  Case 0
    'GR PAGA A PRESTADOR
    'objUserGRArqCons.icTipoGR = tpIcTipoGR_Prest
    'objUserGRArqCons.strGR = "GR Paga a prestador"
    objUserGRPreCons.Show vbModal
'''  Case 1
'''    frmUserProntuarioLis.IcProntuario = tpIcProntuario_Pac
'''    frmUserProntuarioLis.Show vbModal
  End Select
  Set objUserGRPreCons = Nothing
  AmpN
End Sub

Private Sub mnuCaixa_Click(Index As Integer)
  On Error GoTo trata
  AmpS
  Select Case Index
  Case 0
    frmUserGRCons.Show vbModal
  Case 2
    frmUserProntuarioLis.IcProntuario = tpIcProntuario_Pac
    frmUserProntuarioLis.Show vbModal
  Case 4: frmUserTurnoInc.Show vbModal
  End Select
  AmpN
  Exit Sub
trata:
    TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub mnuCancFinan_Click(Index As Integer)
  Dim objUserGRPagamentoLis As SisMed.frmUserGRPagamentoLis
  AmpS
  Set objUserGRPagamentoLis = New SisMed.frmUserGRPagamentoLis
  Select Case Index
  Case 0
    'CANCELADA PONTUAL
    objUserGRPagamentoLis.icTipoGR = tpIcTipoGR_CancPont
    objUserGRPagamentoLis.strGR = "Cancelamento pontual"
    objUserGRPagamentoLis.Show vbModal
  Case 1
    'CANCELADA AUTOMATICA
    objUserGRPagamentoLis.icTipoGR = tpIcTipoGR_CancAut
    objUserGRPagamentoLis.strGR = "Cancelamento automático"
    objUserGRPagamentoLis.Show vbModal
  End Select
  Set objUserGRPagamentoLis = Nothing
  AmpN
End Sub

Private Sub mnuConcFinan_Click(Index As Integer)
  'AmpS
  Select Case Index
  Case 0: SisMed.frmUserGRFinancCons.Show vbModal
  End Select
  'AmpN
End Sub

Private Sub mnuDirConfig_Click(Index As Integer)
  'AmpS
  Select Case Index
  Case 0: RegistrarChave
  Case 1: frmUserConfiguracao.Show vbModal
  End Select
  'AmpN
End Sub

Private Sub mnuDirDespesa_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserGrupoDespesaLis.Show vbModal
  Case 1: frmUserFormaPgtoLis.Show vbModal
  Case 2: frmUserLivroLis.Show vbModal
  Case 4
    frmUserDespesaCtaLis.strTipo = "A" 'Administração
    frmUserDespesaCtaLis.strTipoCtaPagas = "N"
    frmUserDespesaCtaLis.Show vbModal
'''  Case 5
'''    frmUserDespesaCtaLis.strTipo = "A" 'Administração
'''    frmUserDespesaCtaLis.strTipoCtaPagas = "S"
'''    frmUserDespesaCtaLis.Show vbModal
'''  Case 6: frmUserSaldoLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuDirFinanc_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserBancoLis.Show vbModal
  Case 1: frmUserCartaoLis.Show vbModal
  Case 2: frmUserCartaoDebitoLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuDirGR_Click(Index As Integer)
  On Error GoTo trata
  Dim objForm As SisMed.frmUserContaCorrente
  AmpS
  Select Case Index
  Case 0
    frmUserGRCons.Show vbModal
  Case 2
    frmUserProntuarioLis.IcProntuario = tpIcProntuario_Pac
    frmUserProntuarioLis.Show vbModal
  Case 3
    frmUserProntuarioLis.IcProntuario = tpIcProntuario_Prest
    frmUserProntuarioLis.Show vbModal
  Case 5: frmUserProcedimentoLis.Show vbModal
  Case 6: frmUserTipoProcedimentoLis.Show vbModal
  Case 7: frmUserEspecialidadeLis.Show vbModal
  End Select
  AmpN
  Exit Sub
trata:
    TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub mnuDirSala_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserSalaLis.Show vbModal
  Case 2: frmUserPredioLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuDirTurno_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserTurnoInc.Show vbModal
  Case 2: frmUserPeriodoLis.Show vbModal
  Case 3: frmUserDiasDaSemanaLis.Show vbModal
  End Select
  AmpN
End Sub


Private Sub mnuFinanceiro_Click(Index As Integer)
  Dim objUserGRFinCons As SisMed.frmUserGRFinCons
  AmpS
  Set objUserGRFinCons = New SisMed.frmUserGRFinCons
  Select Case Index
  Case 0
    'GR PAGA A PRESTADOR
    'objUserGRArqCons.icTipoGR = tpIcTipoGR_Prest
    'objUserGRArqCons.strGR = "GR Paga a prestador"
    objUserGRFinCons.Show vbModal
'''  Case 1
'''    frmUserProntuarioLis.IcProntuario = tpIcProntuario_Pac
'''    frmUserProntuarioLis.Show vbModal
  End Select
  Set objUserGRFinCons = Nothing
  AmpN

End Sub

Private Sub mnuGerDespesa_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserGrupoDespesaLis.Show vbModal
  Case 1: frmUserFormaPgtoLis.Show vbModal
  Case 2: frmUserLivroLis.Show vbModal
  Case 4
    frmUserDespesaCtaLis.strTipo = "A" 'Administração
    frmUserDespesaCtaLis.strTipoCtaPagas = "N"
    frmUserDespesaCtaLis.Show vbModal
'''  Case 5
'''    frmUserDespesaCtaLis.strTipo = "A" 'Administração
'''    frmUserDespesaCtaLis.strTipoCtaPagas = "S"
'''    frmUserDespesaCtaLis.Show vbModal
'''  Case 6: frmUserSaldoLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuGerFinanc_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserBancoLis.Show vbModal
  Case 1: frmUserCartaoLis.Show vbModal
  Case 2: frmUserCartaoDebitoLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuGerGR_Click(Index As Integer)
  On Error GoTo trata
  AmpS
  Select Case Index
  Case 0
    frmUserGRCons.Show vbModal
  Case 2
    frmUserProntuarioLis.IcProntuario = tpIcProntuario_Pac
    frmUserProntuarioLis.Show vbModal
  Case 3
    frmUserProntuarioLis.IcProntuario = tpIcProntuario_Prest
    frmUserProntuarioLis.Show vbModal
  Case 5: frmUserProcedimentoLis.Show vbModal
  Case 6: frmUserTipoProcedimentoLis.Show vbModal
  Case 7: frmUserEspecialidadeLis.Show vbModal
  End Select
  AmpN
  Exit Sub
trata:
    TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub mnuGerSala_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserSalaLis.Show vbModal
  Case 2: frmUserPredioLis.Show vbModal
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

Private Sub mnuLaboratorio_Click(Index As Integer)
  On Error GoTo trata
  AmpS
  Select Case Index
  Case 0
    frmUserGRCons.Show vbModal
  Case 2
    frmUserProntuarioLis.IcProntuario = tpIcProntuario_Pac
    frmUserProntuarioLis.Show vbModal
  Case 4: frmUserTurnoInc.Show vbModal
  End Select
  AmpN
  Exit Sub
trata:
    TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub mnuPagarPrest_Click(Index As Integer)
  Dim objUserGRPagamentoLis As SisMed.frmUserGRPagamentoLis
  AmpS
  Set objUserGRPagamentoLis = New SisMed.frmUserGRPagamentoLis
  Select Case Index
  Case 0
    'GR PAGA A PRESTADOR
    objUserGRPagamentoLis.icTipoGR = tpIcTipoGR_Prest
    objUserGRPagamentoLis.strGR = "GR Paga a prestador"
    objUserGRPagamentoLis.Show vbModal
  Case 1
    'GR PAGA A TEC RX
    objUserGRPagamentoLis.icTipoGR = tpIcTipoGR_TecRX
    objUserGRPagamentoLis.strGR = "GR Paga a tec. RX"
    objUserGRPagamentoLis.Show vbModal
  Case 2
    'GR PAGA A dono de RX
    objUserGRPagamentoLis.icTipoGR = tpIcTipoGR_DonoRX
    objUserGRPagamentoLis.strGR = "GR Paga a dono de RX"
    objUserGRPagamentoLis.Show vbModal
  Case 3
    'GR PAGA A dono de Ultrason
    objUserGRPagamentoLis.icTipoGR = tpIcTipoGR_DonoUltra
    objUserGRPagamentoLis.strGR = "GR Paga a dono de Ultrason"
    objUserGRPagamentoLis.Show vbModal
  End Select
  Set objUserGRPagamentoLis = Nothing
  AmpN
End Sub


Private Sub mnuRelFinanc_Click(Index As Integer)
  On Error GoTo trata
  AmpS
  Select Case Index
  Case 0: frmUserRelPgtoPrest.Show vbModal
  Case 1: frmUserRelPgtoPrestDC.Show vbModal
  Case 2: frmUserRelPgtoPrestCart.Show vbModal
  Case 3: frmUserRelPgtoPrestCartCons.Show vbModal
  Case 4: frmUserRelProc.Show vbModal
  Case 5: frmUserRelPrestEsp.Show vbModal
  Case 6: frmUserRelAtendCaixa.Show vbModal
  Case 7: frmUserRelContas.Show vbModal
  Case 8: frmUserRelResumoDespesas.Show vbModal
  Case 9: frmUserRelGRAtend.Show vbModal
  Case 10: frmUserRelGRNPagaPrest.Show vbModal
  Case 11: frmUserRelGRCanc.Show vbModal
  Case 12: frmUserRelGRCancFunc.Show vbModal
  Case 13: frmUserRelNovoPaciente.Show vbModal
  Case 14: frmUserRelNovoPacienteCons.Show vbModal
  End Select
  AmpN
  Exit Sub
trata:
    TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub mnuRelGerenc_Click(Index As Integer)
  On Error GoTo trata
  AmpS
  Select Case Index
  Case 0: frmUserRelGerencial.Show vbModal
  Case 1: frmUserRelDiaReceita.Show vbModal
  Case 2
    frmUserRelBalancoInc.Status = tpStatus_Consultar
    frmUserRelBalancoInc.Show vbModal
  
  Case 3: frmUserRelPgtoPrestProc.Show vbModal
  Case 4: frmUserRelReceita.Show vbModal
  Case 5: frmUserRelReceitaEspec.Show vbModal
  Case 6: frmUserRelPgtoPrestProcPg.Show vbModal
  Case 7: frmUserRelCompReceita.Show vbModal
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
