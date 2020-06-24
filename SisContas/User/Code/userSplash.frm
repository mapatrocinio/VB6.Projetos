VERSION 5.00
Begin VB.Form frmUserSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Start"
   ClientHeight    =   3735
   ClientLeft      =   1980
   ClientTop       =   360
   ClientWidth     =   6375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "userSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3735
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSistema 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Carregando Sistema"
      Enabled         =   0   'False
      Height          =   225
      Left            =   3570
      TabIndex        =   4
      Top             =   900
      Width           =   2010
   End
   Begin VB.CheckBox chkSenha 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Obtendo Senha"
      Enabled         =   0   'False
      Height          =   210
      Left            =   480
      TabIndex        =   3
      Top             =   1350
      Width           =   1995
   End
   Begin VB.CheckBox chkBanco 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lendo Banco de Dados"
      Enabled         =   0   'False
      Height          =   210
      Left            =   480
      TabIndex        =   2
      Top             =   1110
      Width           =   2040
   End
   Begin VB.CheckBox chkConexao 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Carregando Conexão"
      Enabled         =   0   'False
      Height          =   210
      Left            =   480
      TabIndex        =   1
      Top             =   885
      Width           =   2100
   End
   Begin VB.Timer Timer2 
      Left            =   3720
      Top             =   1005
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   6240
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6345
   End
End
Attribute VB_Name = "frmUserSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub LerFiguras()
  On Error Resume Next
  '
  Me.Picture = LoadPicture(gsBMPPath & "Xa_2.jpg")
  '
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim objRegister As busSisContas.clsRegistro
  lblTitulo.Caption = TITULOSISTEMA
  chkConexao.Value = 1
  chkBanco.Value = 1
  'Inicializar variáveis do register
  Set objRegister = New busSisContas.clsRegistro
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
  Set objRegister = Nothing
  
  chkSenha.Value = 1
  '
  LerFiguras
  '
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error GoTo trata
  frmMDI.stbPrinc.Panels(1).Text = gsNomeUsu
  frmMDI.stbPrinc.Panels(2).Text = gsNivel
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub Timer2_Timer()
  On Error Resume Next
  Dim RetVal As Variant
  '
  'HabServDesp
  'RetVal = Shell(gsAppPath & "AUTO.EXE", vbMinimizedNoFocus)    ' Run Calculator.
  'Captura_Config
  'HabServDesp
  '
  'Fazer Verificação para estoqusta para saber se quantidade mínima atingiu o estoque
  If gsNivel = "EST" Then
    If Verifica_Estoque Then
      'frmVisualizarEstoque.QuemChamou = 1
      'frmVisualizarEstoque.Show
      frmUserEstoqueViz.QuemChamou = 1
      frmUserEstoqueViz.Show vbModal
    
    End If
  End If
  '
'  If gbTrabComImpFiscal Then
'    'SOMENTE ABRE A IMPRESSORA FISCAL CASO SEJA SENHA RECEPÇÃO
'    If gsNivel = "REC" Then
'      'INIFIO DO TRATAMENTO PARA IMPRESSÃO FISCAL
'      '-----------------------------------------
'      LocalRetorno = LeParametrosIni("Sistema", "Retorno")
'      LocalRetorno = Left(LocalRetorno, 1)
'
'      Retorno = Bematech_FI_AbrePortaSerial
'      VerificaRetornoImpressora "", "", "BemaFI32"
'
'      If Retorno = -4 Or Retorno = -5 Then
'        MsgBox "A impressora Fiscal não foi inicializada. O programa seguirá SEM UTILIZA-LA.", vbExclamation, TITULOSISTEMA
'        gbTrabComImpFiscal = False
'      End If
'      '-----------------------------------------
'      'TÉRMINO DO TRATAMENTO PARA IMPRESSÃO FISCAL
'    Else
'      MsgBox "A impressora fiscal somente será inicializada para Nível de RECEPÇÃO. A impressora fiscal NÃO FOI INICIALIZADA PARA SEU NÍVEL DE SENHA, portanto RECEBIMENTOS NESTA ESTAÇÃO NÃO SERÃO COMPUTADOS NA IMPRESSORA FISCAL.", vbExclamation, TITULOSISTEMA
'      gbTrabComImpFiscal = False
'    End If
'  End If
  'TratarTemporada
  Unload Me
  frmMDI.Show
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  End
End Sub


