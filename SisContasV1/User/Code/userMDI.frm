VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8970
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Width           =   8970
      _ExtentX        =   15822
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
            TextSave        =   "25/08/2003"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "16:10"
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
   Begin VB.Menu snuAdministracao 
      Caption         =   "&Administração"
      Index           =   0
      Begin VB.Menu snuDespesas 
         Caption         =   "Despesas"
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
         Begin VB.Menu mnuDespesas 
            Caption         =   "&Saldo"
            Index           =   6
         End
      End
      Begin VB.Menu snuMovimentacoes 
         Caption         =   "Movimentações"
         Begin VB.Menu mnuMovimentacoes 
            Caption         =   "&Contas"
            Index           =   0
         End
         Begin VB.Menu mnuMovimentacoes 
            Caption         =   "Zerar Contas"
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
   End
   Begin VB.Menu snuRelatorio 
      Caption         =   "&Relatórios"
      Index           =   0
      Begin VB.Menu snuRelFinanc 
         Caption         =   "&Financeiros"
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Demonstrativo de Contas"
            Index           =   0
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Balanço Sismotel"
            Index           =   1
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Balanço Movimento"
            Index           =   2
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "Demonstrativo &Resumo Geral de Despesas"
            Index           =   3
         End
         Begin VB.Menu mnuRelFinanc 
            Caption         =   "&Listagem de Grupo / Sub Grupo"
            Index           =   4
         End
      End
      Begin VB.Menu snuContas 
         Caption         =   "&Contas"
         Begin VB.Menu mnuContas 
            Caption         =   "&Extrato de Movimentação"
            Index           =   0
         End
         Begin VB.Menu mnuContas 
            Caption         =   "Extrato de &Saldos"
            Index           =   1
         End
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bPrimeiraVez As Boolean

Public Sub LerFiguras()
   '
   Me.Icon = LoadPicture(gsIconsPath & "Logo.ico")
   '
End Sub

Private Sub MDIForm_Activate()
  If App.PrevInstance Then
    MsgBox "Aplicativo já está rodando!", vbExclamation, TITULOSISTEMA
    End
  End If
  
  
  
'  'Eugenio, para tirar a proteção, comente o código abaixo até antes de End Sub
'  'Depois vá em Project/References e desmarque as referencias para Protec
'  '---------------------------------------------------------------
'  '----------------
'  'Proteção do sistema
'  '----------------
'  Dim clsProtec As clsProtec
'  Set clsProtec = New clsProtec
'  '----------------
'  'Verifica Proteção do sistema
'  '-------------------------
'  'Valida primeira vez que entrou no sistema
'  If Not clsProtec.Valida_Primeira_Vez(gsBDadosPath & nomeBDados, App.Path) Then
'    End
'    Exit Sub
'  End If
'  'Válida Equipamento
'  If Not clsProtec.Valida_Estacao(gsBDadosPath & nomeBDados) Then
'    End
'    Exit Sub
'  End If
'  Set clsProtec = Nothing
'  '----------------
'  Set clsProtec = New clsProtec
'  'Valida se sistema expirou
'  If Not clsProtec.Valida_Chave(gsBDadosPath & nomeBDados, IIf(bPrimeiraVez, "S", "N"), gsNivel) Then
'    End
'    Exit Sub
'  End If
'  'Atualizar data Atual do sistema
'  clsProtec.Atualiza_Chave_Data_Atual gsBDadosPath & nomeBDados
'  'Mata o arquivo fisicamene
'  clsProtec.Trata_Arquivo_Fisico App.Path
'  Set clsProtec = Nothing
'  '-----------------
'  '------------ FIM
'  '----------------
'


  bPrimeiraVez = False
  '---------------------------------------------------------------
End Sub

Private Sub MDIForm_Load()
  On Error Resume Next
  bPrimeiraVez = True
  AmpS
  '
  ConnectRpt = "ODBC;PWD=SHOGUM2806;DATABASE=" & gsBDadosPath & "SisMotel.MDB"
  Me.Caption = TITULOSISTEMA & " - " & gsNomeEmpresa
  '
  LerFiguras
  '
  If Len(Trim(gsBMP)) <> 0 Then
    If Dir(gsBMP) <> "" Then
      Me.Picture = LoadPicture(gsBMP)
    End If
  End If
  AmpN
  '
  'Monta_Menu 1
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

Private Sub mnuArquivo_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0
    If frmMDI.mnuArquivo(0).Caption = "&Desconectar" Then
      frmMDI.mnuArquivo(0).Caption = "&Conectar"
      'Monta_Menu 0
      '
      'Captura configurações do Usuário
      gsNomeUsu = ""
      gsNivel = ""
      '
      frmMDI.stbPrinc.Panels(1).Text = gsNomeUsu
      frmMDI.stbPrinc.Panels(2).Text = gsNivel
      '
      'Captura_Config
      '
    Else
      frmLogin.QuemChamou = 1
      frmLogin.Show vbModal
      '
      frmMDI.stbPrinc.Panels(1).Text = gsNomeUsu
      frmMDI.stbPrinc.Panels(2).Text = gsNivel
      '
      'Captura_Config
      '
    End If
  '
  Case 2: 'frmUsuarios.Show
  Case 4: 'frmPapel.Show
  Case 6: Unload Me
  End Select
  AmpN
End Sub

Private Sub mnuContas_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserRelExtrato.Show vbModal
  Case 1: frmUserRelSaldo.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuDespesas_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserGrupoDespesaLis.Show vbModal
  Case 1: frmUserFormaPgtoLis.Show vbModal
  Case 2: frmUserLivroLis.Show vbModal
  Case 4
    frmUserDespesaLis.strTipo = "A" 'Administração
    frmUserDespesaLis.strTipoCtaPagas = "N"
    frmUserDespesaLis.Show vbModal
  Case 5
    frmUserDespesaLis.strTipo = "A" 'Administração
    frmUserDespesaLis.strTipoCtaPagas = "S"
    frmUserDespesaLis.Show vbModal
  Case 6: frmUserSaldoLis.Show vbModal
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

Private Sub mnuRelFinanc_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserRelContas.Show vbModal
  Case 1: frmUserRelBalanco.Show vbModal
  Case 2: frmUserRelBalancoMov.Show vbModal
  Case 3: frmUserRelResumoDespesas.Show vbModal
  Case 4: frmUserRelGrupoSubGrupo.Show vbModal
  End Select
  AmpN
End Sub

Private Sub snuSobre_Click()
  AmpS
  frmAbout.Show
  AmpN
End Sub

