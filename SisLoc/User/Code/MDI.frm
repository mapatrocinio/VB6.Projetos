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
            TextSave        =   "24/7/2010"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   2
            TextSave        =   "13:14"
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
         Caption         =   "&Funcion�rio"
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
      Caption         =   "&Diretoria"
      Index           =   0
      Begin VB.Menu mnuGerencia 
         Caption         =   "&NFSR"
         Index           =   0
         Begin VB.Menu mnuGerNF 
            Caption         =   "&NFSR"
            Index           =   0
         End
         Begin VB.Menu mnuGerNF 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuGerNF 
            Caption         =   "&Empresa/Contrato/Obra"
            Index           =   2
         End
         Begin VB.Menu mnuGerNF 
            Caption         =   "&BM"
            Index           =   3
         End
         Begin VB.Menu mnuGerNF 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuGerNF 
            Caption         =   "&Tipo de Empresa"
            Index           =   5
         End
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "&Estoque"
         Index           =   1
         Begin VB.Menu mnuGerEstoque 
            Caption         =   "&Estoque"
            Index           =   0
         End
         Begin VB.Menu mnuGerEstoque 
            Caption         =   "&Unidade"
            Index           =   1
         End
         Begin VB.Menu mnuGerEstoque 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuGerEstoque 
            Caption         =   "&Entrada Material"
            Index           =   3
         End
         Begin VB.Menu mnuGerEstoque 
            Caption         =   "&Documento"
            Index           =   4
         End
      End
      Begin VB.Menu mnuGerencia 
         Caption         =   "&Configura��o"
         Index           =   2
         Begin VB.Menu mnuGerConfig 
            Caption         =   "&Registrar chave"
            Index           =   0
         End
         Begin VB.Menu mnuGerConfig 
            Caption         =   "&Configura��o"
            Index           =   1
         End
      End
   End
   Begin VB.Menu snuFinanceiro 
      Caption         =   "&Financeiro"
      Index           =   0
      Begin VB.Menu mnuFinanceiro 
         Caption         =   "&BM"
         Index           =   0
      End
   End
   Begin VB.Menu snuCaixa 
      Caption         =   "&Caixa"
      Index           =   0
      Begin VB.Menu mnuCaixa 
         Caption         =   "&NFSR"
         Index           =   0
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
  Dim objGeral    As busSisLoc.clsGeral
  Dim objProtec   As busSisLoc.clsProtec
  '
  If App.PrevInstance Then
    MsgBox "Aplicativo j� est� rodando!", vbExclamation, TITULOSISTEMA
    End
  End If
  If blnPrimeiraVez Then
    '-----------------
    '------------ INICIO
    '----------------
    'Eugenio, para tirar a prote��o, comente o c�digo abaixo at� antes de End Sub
    'Depois v� em Project/References e desmarque as referencias para Protec
    '---------------------------------------------------------------
    '----------------
    'Prote��o do sistema
    '----------------
    Set objProtec = New busSisLoc.clsProtec
    Set objGeral = New busSisLoc.clsGeral
    '----------------
    'Verifica Prote��o do sistema
    '-------------------------
    'Valida primeira vez que entrou no sistema
    If Not objProtec.Valida_Primeira_Vez(objGeral.ObterConnectionString, App.Path) Then
      Set objProtec = Nothing
      Set objGeral = Nothing
      End
      Exit Sub
    End If
    'V�lida Equipamento
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
'''      TratarErroPrevisto "Acabou a validade desta c�pia do sistema, contacte o suporte para adquirir uma nova vers�o.", "[frmMDI_Activqte]"
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
    Case gsFinanceiro
      'BM
      frmUserBMLis.Show vbModal
    Case gsCaixa
      'NFSR
      frmUserNFCons.Show vbModal
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
    '  MsgBox "Voc� n�o tem autoriza��o para sair do sistema. Para efetuar essa opera��o, v� em arquivo/Desconectar, depois v� em arquivo/Conectar e chame seu gerente/Diretor para entrar com a senha e sair do sistema.", vbExclamation, TITULOSISTEMA
    '  Cancel = True
    'Else
      CapturaParametrosRegistro 3
      End
    'End If
End Sub
'''
'''Private Sub tmrProtecao_Timer()
'''  On Error GoTo trata
'''  Dim objProtec As busSisLoc.clsProtec
'''  Dim objGeral As busSisLoc.clsGeral
'''  Set objProtec = New busSisLoc.clsProtec
'''  Set objGeral = New busSisLoc.clsGeral
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
      'Captura configura��es do Usu�rio
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



Private Sub mnuCaixa_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserNFCons.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuFinanceiro_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserBMLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuGerConfig_Click(Index As Integer)
  'AmpS
  Select Case Index
  'Case 0: RegistrarChave
  Case 1: frmUserConfiguracao.Show vbModal
  End Select
  'AmpN
End Sub

Private Sub mnuGerEstoque_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserEstoqueLis.Show vbModal
  Case 1: frmUserUnidadeLis.Show vbModal
  Case 3: frmUserEntradaMaterialLis.Show vbModal
  Case 4: frmUserDocumentoLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub mnuGerNF_Click(Index As Integer)
  AmpS
  Select Case Index
  Case 0: frmUserNFCons.Show vbModal
  Case 2: frmUserEmpresaLis.Show vbModal
  Case 3: frmUserBMLis.Show vbModal
  Case 5: frmUserTipoEmpresaLis.Show vbModal
  End Select
  AmpN
End Sub

Private Sub snuSobre_Click()
  AmpS
  frmUserAbout.Show
  AmpN
End Sub
