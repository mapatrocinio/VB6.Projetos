VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfigCadastro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurações do Sistema - Módulo de Dados Cadastrais"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5865
      Left            =   8520
      ScaleHeight     =   5865
      ScaleWidth      =   1860
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   90
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3660
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5595
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   150
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9869
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados Cadastrais"
      TabPicture(0)   =   "userConfigCadastro.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraConfiguracao(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraConfiguracao 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Configuração"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Index           =   0
         Left            =   150
         TabIndex        =   16
         Top             =   390
         Width           =   7905
         Begin VB.TextBox txtTitulo 
            Height          =   288
            Left            =   1200
            MaxLength       =   255
            TabIndex        =   9
            Top             =   2400
            Width           =   6492
         End
         Begin VB.CommandButton cmdLimparBase 
            Caption         =   "Limpar Base de Dados"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   2820
            Width           =   2295
         End
         Begin VB.TextBox txtEmpresa 
            Height          =   288
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   0
            Top             =   270
            Width           =   6492
         End
         Begin VB.TextBox txtEndereco 
            Height          =   288
            Left            =   1200
            MaxLength       =   100
            TabIndex        =   3
            Top             =   960
            Width           =   6492
         End
         Begin VB.TextBox txtBairro 
            Height          =   288
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1320
            Width           =   6492
         End
         Begin VB.TextBox txtCidade 
            Height          =   288
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   5
            Top             =   1680
            Width           =   6492
         End
         Begin VB.TextBox txtEstado 
            Height          =   288
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   6
            Top             =   2040
            Width           =   612
         End
         Begin VB.TextBox txtTelefone 
            Height          =   288
            Left            =   4920
            MaxLength       =   30
            TabIndex        =   8
            Top             =   2040
            Width           =   2775
         End
         Begin VB.TextBox txtCnpj 
            Height          =   288
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   1
            Top             =   600
            Width           =   2355
         End
         Begin VB.TextBox txtInscrMunicipal 
            Height          =   288
            Left            =   5370
            MaxLength       =   20
            TabIndex        =   2
            Top             =   600
            Width           =   2325
         End
         Begin VB.TextBox txtCep 
            Height          =   288
            Left            =   2670
            MaxLength       =   10
            TabIndex        =   7
            Top             =   2040
            Width           =   1275
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Título"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   26
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Empresa"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   270
            Width           =   975
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Endereço"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Bairro"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Cidade"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Estado"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   21
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Telefone"
            Height          =   255
            Index           =   5
            Left            =   3960
            TabIndex        =   20
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Cnpj"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Inscr. Estad."
            Height          =   255
            Index           =   13
            Left            =   4230
            TabIndex        =   18
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Cep"
            Height          =   255
            Index           =   14
            Left            =   1530
            TabIndex        =   17
            Top             =   2070
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frmConfigCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                     As tpStatus
Public lngCONFIGID          As Long
Public bRetorno                   As Boolean
Public bFechar                    As Boolean
Private blnPrimeiraVez            As Boolean



Private Sub cmdCancelar_Click()
  bFechar = True
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdLimparBase_Click()
  On Error GoTo trata
  Dim objGeral As busSisMetal.clsGeral
  Dim strSql As String
  '
  If MsgBox("ATENÇÃO: Esta operação limpará toda a base de dados. Tem certeza de que deseja continuar?", vbYesNo, TITULOSISTEMA) = vbNo Then Exit Sub
  If MsgBox("ATENÇÃO: Esta operação limpará toda a base de dados. Tem certeza de que deseja continuar?", vbYesNo, TITULOSISTEMA) = vbNo Then Exit Sub
  '
  Set objGeral = New busSisMetal.clsGeral
  
  'TABELAS DIVERSAS
  '
  strSql = "DELETE FROM LOG_ESTOQUE;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM SEQUENCIAL;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  'strSql = "DELETE FROM CHAVE;"
  'objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM LOG_UNIDADE;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM REL_CAMPVENDA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM REL_DEMOOCUPRESERVA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM CARTAOPROMOCIONAL;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM FXCARTAOPROMOCIONAL;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM CLIENTECARTAOPROMOCIONAL;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  
  
  strSql = "DELETE FROM ALERTACLIENTE;"
  objGeral.ExecutarSQLAtualizacao strSql
  '
  '
  'CONTA CORRENTE
  '
  strSql = "DELETE FROM PARCELA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM BAIXA_PENHOR;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM CONTACORRENTE;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  'ESTOQUE
  '
  strSql = "DELETE FROM LOG_ESTOQUE;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_ENTRADAMATERIAL;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM ENTRADAMATERIAL;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_RETORNOREQUISICAO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM RETORNOREQUISICAO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_REQUISICAOMATERIAL;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM REQUISICAOMATERIAL;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_VENDACARD;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  
  
  strSql = "DELETE FROM VENDA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_CARDESTINTER;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_PEDIDOCARD;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM PEDIDO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_RESPRESERVA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM CARDAPIO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM CARDAPIO_RESUMO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_TRANSFESTINTER;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TRANSFESTINTER;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_GRUPOESTAPTO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_RETORNOESTOQUE;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_RETORNOESTOQUEINTERMEDIARIO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM RETORNOESTOQUE;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_BAIXAESTOQUE;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_BAIXAESTOQUEINTERMEDIARIO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM BAIXAESTOQUE;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_DEPOSITOESTINTER;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM DEPOSITO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_GRUPOESTESTINTER;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM ESTOQUEINTERMEDIARIO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM ESTOQUE;"
  objGeral.ExecutarSQLAtualizacao strSql
  '
  'DESPESA
  '
  strSql = "DELETE FROM PENHOR;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM SALDO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM DESPESA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM LIVRO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM SUBGRUPODESPESA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM GRUPODESPESA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM MOVIMENTACAO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM CONTA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TIPOCONTA;"
  objGeral.ExecutarSQLAtualizacao strSql
  '
  'FUNCIONARIO
  '
  strSql = "DELETE FROM FUNCIONARIO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM CONTROLEACESSO WHERE USUARIO <> 'EUGENIO' AND USUARIO <> 'MIGUEL';"
  objGeral.ExecutarSQLAtualizacao strSql
  
  'strSql = "DELETE FROM FUNCAO;"
  'objGeral.ExecutarSQLAtualizacao strSql
  '
  'TEMPORADA/PACOTE
  '
  strSql = "DELETE FROM TAB_TEMPGRUPOPERIODO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TEMPORADA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_PACGRUPOPERIODO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  
  strSql = "DELETE FROM TAB_TEMPGRUPOPERIODO;"
  objGeral.ExecutarSQLAtualizacao strSql
  '
  'HOTEL/RESERVA
  '
  strSql = "DELETE FROM TAB_RESPRESERVA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_RESPLOCACAO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  'strSql = "DELETE FROM RESPONSABILIDADE;"
  'objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM VIAGEM;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  'strSql = "DELETE FROM MEIOTRANSPORTE;"
  'objGeral.ExecutarSQLAtualizacao strSql
  
  'strSql = "DELETE FROM MOTIVOVIAGEM;"
  'objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_LOCASSOC;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM ASSOCIACAO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_FICHACLIELOC;"
  objGeral.ExecutarSQLAtualizacao strSql
  '
  'APÓS DELETAR LOCACAO, DELETA DEMAIS TABELAS DE HOTEL/RESERVA
  '
  'LOCACAO
  '
  strSql = "DELETE FROM TAB_EXTRAUNIDADE;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM EXTRAUNIDADE;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM DESPESA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM SANGRIA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM CHEQUE;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM CLIENTE;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM CHQSDEVOLVIDOS;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM LOG_UNIDADE;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM DESPERTADOR;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TELEFONEMA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_EXTRA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM EXTRA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM LEMBRETE;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM CORTESIA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_GRUPO_TIPOCORTESIA_DIASEMANA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_TIPOCORTESIA_DIASEMANA;"
  objGeral.ExecutarSQLAtualizacao strSql

  'strSql = "DELETE FROM TIPO_CORTESIA;"
  'objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_FICHACLIELOC;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM VIAGEM;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_RESPLOCACAO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM PARCELA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM CONTACORRENTE;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_LOCASSOC;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM ASSOCIACAO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_EXTRA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM EXTRA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM ENTRADASAIDA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM LOCACAO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  
  strSql = "DELETE FROM CAMAREIRA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  'strSql = "DELETE FROM CARTAO;"
  'objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM GARCOM;"
  objGeral.ExecutarSQLAtualizacao strSql
  '
  'DEMAIS DADOS DE HOTEL/RESERVA E LOCAÇÃO
  '
  'HOTEL/RESERVA
      
  strSql = "DELETE FROM RESERVA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM PACOTE;"
  objGeral.ExecutarSQLAtualizacao strSql
      
  strSql = "DELETE FROM TAB_RESPRESERVA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
      
  'strSql = "DELETE FROM TIPOPAGAMENTO;"
  'objGeral.ExecutarSQLAtualizacao strSql
      
  'strSql = "DELETE FROM GARANTIA;"
  'objGeral.ExecutarSQLAtualizacao strSql
      
  strSql = "DELETE FROM PACOTE;"
  objGeral.ExecutarSQLAtualizacao strSql
      
  strSql = "DELETE FROM FICHACLIENTE;"
  objGeral.ExecutarSQLAtualizacao strSql
    
  'strSql = "DELETE FROM TIPODOCUMENTO;"
  'objGeral.ExecutarSQLAtualizacao strSql
    
  strSql = "DELETE FROM EMPRESA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  'strSql = "DELETE FROM TIPOEMPRESA;"
  'objGeral.ExecutarSQLAtualizacao strSql
      
  'CONFIGURACAO
  
  strSql = "DELETE FROM TAB_GRUPOPERIODO_CONFIG;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM TAB_APTO_CONFIG;"
  objGeral.ExecutarSQLAtualizacao strSql
    
  strSql = "DELETE FROM TAB_FAIXA_CONFIG;"
  objGeral.ExecutarSQLAtualizacao strSql
    
  strSql = "DELETE FROM GRUPOPERIODO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM FXGRUPOPERIODO;"
  objGeral.ExecutarSQLAtualizacao strSql
    
  'LOCAÇÃO
  
  
  strSql = "DELETE FROM INTERDICAO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM APARTAMENTO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  strSql = "DELETE FROM GRUPO;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  '
  'TURNO
  '
  strSql = "DELETE FROM OCORRENCIA;"
  objGeral.ExecutarSQLAtualizacao strSql
  
  
  strSql = "DELETE FROM TURNO;"
  objGeral.ExecutarSQLAtualizacao strSql
  '
  'FXGRUPOPERIODO
  '
  strSql = "DELETE FROM FXGRUPOPERIODO;"
  objGeral.ExecutarSQLAtualizacao strSql
  '
  MsgBox "Operação realizada com sucesso !", vbExclamation, TITULOSISTEMA
  '
  Set objGeral = Nothing
  '
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim strMsgErro              As String
  Dim objRs                   As ADODB.Recordset
  Dim objConfiguracao         As busSisMetal.clsConfiguracao
  Dim objGer                  As busSisMetal.clsGeral

  '
  If Not ValidaCampos Then Exit Sub
  '
  Set objConfiguracao = New busSisMetal.clsConfiguracao
  '
  If Status = tpStatus_Alterar Then
    'Código para alteração
    '
    objConfiguracao.AlterarConfiguracaoCadastro lngCONFIGID, _
                                                txtEmpresa.Text, _
                                                txtCnpj.Text, _
                                                txtInscrMunicipal.Text, _
                                                txtEndereco.Text, _
                                                txtBairro.Text, _
                                                txtCidade.Text, _
                                                txtEstado.Text, _
                                                txtCep.Text, _
                                                txtTelefone.Text, _
                                                txtTitulo.Text

    '
    Captura_Config
    '
    bRetorno = True
  ElseIf Status = tpStatus_Incluir Then
  End If
  Set objConfiguracao = Nothing
  bFechar = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg          As String
  Dim blnSetarFoco    As Boolean
  '
  blnSetarFoco = True
  If Not Valida_String(txtTitulo, TpObrigatorio, blnSetarFoco) Then
    strMsg = "Preencher o título da configuração"
    'blnSetarFoco = False
    'tabDetalhes.Tab = 1
    'Pintar_Controle mskQtdCadastro, tpCorContr_Erro
    'SetarFoco mskQtdCadastro
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserConfigCadastro.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    tabDetalhes.Tab = 0
    tabDetalhes_Click 0
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserConfigCadastro.Form_Activate]"
End Sub

Private Sub LimparCampos()
  On Error GoTo trata
  'Configuração de Cadastro
  LimparCampoTexto txtEmpresa
  LimparCampoTexto txtCnpj
  LimparCampoTexto txtInscrMunicipal
  LimparCampoTexto txtEndereco
  LimparCampoTexto txtBairro
  LimparCampoTexto txtCidade
  LimparCampoTexto txtEstado
  LimparCampoTexto txtCep
  LimparCampoTexto txtTelefone
  LimparCampoTexto txtTitulo
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserConfigImpressao.LimparCampos]", _
            Err.Description
            
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs             As ADODB.Recordset
  Dim strSql            As String
  Dim objGeral          As busSisMetal.clsGeral
  Dim objConfiguracao   As busSisMetal.clsConfiguracao
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 6345
  Me.Width = 10470
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  'Capturar configurações do sistema
  'Set objGeral = New busSisMetal.clsGeral
  'strSql = "SELECT PKID FROM CONFIGURACAO"
  'Set objRs = objGeral.ExecutarSQL(strSql)
  'If objRs.EOF Then
  '  'Inclusão
  '  Err.Raise 999, , "Não há registro de configuração cadastrado!"
  'Else
  '  'Alteração
  '  Status = tpStatus.tpStatus_Alterar
  '  lngCONFIGID = objRs.Fields("PKID").Value
  'End If
  'objRs.Close
  'Set objRs = Nothing
  'Set objGeral = Nothing
  Status = tpStatus.tpStatus_Alterar
  'Limpar Campos
  LimparCampos
  If gsNivel = gsAdmin Then
    cmdLimparBase.Visible = False
  Else
    cmdLimparBase.Visible = False
  End If
  If Status = tpStatus_Incluir Then
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objConfiguracao = New busSisMetal.clsConfiguracao
    Set objRs = objConfiguracao.ListarConfiguracaoCadastro(lngCONFIGID)
    '
    If Not objRs.EOF Then
      txtEmpresa.Text = objRs.Fields("Empresa").Value & ""
      txtCnpj.Text = objRs.Fields("Cnpj").Value & ""
      txtInscrMunicipal.Text = objRs.Fields("InscrMunicipal").Value & ""
      txtEndereco.Text = objRs.Fields("Endereco").Value & ""
      txtBairro.Text = objRs.Fields("Bairro").Value & ""
      txtCidade.Text = objRs.Fields("Cidade").Value & ""
      txtEstado.Text = objRs.Fields("Estado").Value & ""
      txtCep.Text = objRs.Fields("Cep").Value & ""
      txtTelefone.Text = objRs.Fields("Tel").Value & ""
      txtTitulo.Text = objRs.Fields("Titulo").Value & ""
      '
    End If
    objRs.Close
    Set objRs = Nothing
    Set objConfiguracao = Nothing
  End If

  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not bFechar Then Cancel = True
End Sub

Private Sub tabDetalhes_Click(PreviousTab As Integer)
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    'Desabilitar campos
    fraConfiguracao(0).Enabled = True
    SetarFoco txtEmpresa
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub txtBairro_GotFocus()
  Seleciona_Conteudo_Controle txtBairro
End Sub
Private Sub txtBairro_LostFocus()
  Pintar_Controle txtBairro, tpCorContr_Normal
End Sub

Private Sub txtCep_GotFocus()
  Seleciona_Conteudo_Controle txtCep
End Sub
Private Sub txtCep_LostFocus()
  Pintar_Controle txtCep, tpCorContr_Normal
End Sub

Private Sub txtCidade_GotFocus()
  Seleciona_Conteudo_Controle txtCidade
End Sub
Private Sub txtCidade_LostFocus()
  Pintar_Controle txtCidade, tpCorContr_Normal
End Sub

Private Sub txtCNPJ_GotFocus()
  Seleciona_Conteudo_Controle txtCnpj
End Sub
Private Sub txtCnpj_LostFocus()
  Pintar_Controle txtCnpj, tpCorContr_Normal
End Sub

Private Sub txtEmpresa_GotFocus()
  Seleciona_Conteudo_Controle txtEmpresa
End Sub
Private Sub txtEmpresa_LostFocus()
  Pintar_Controle txtEmpresa, tpCorContr_Normal
End Sub

Private Sub txtEndereco_GotFocus()
  Seleciona_Conteudo_Controle txtEndereco
End Sub
Private Sub txtEndereco_LostFocus()
  Pintar_Controle txtEndereco, tpCorContr_Normal
End Sub

Private Sub txtEstado_GotFocus()
  Seleciona_Conteudo_Controle txtEstado
End Sub
Private Sub txtEstado_LostFocus()
  Pintar_Controle txtEstado, tpCorContr_Normal
End Sub

Private Sub txtInscrMunicipal_GotFocus()
  Seleciona_Conteudo_Controle txtInscrMunicipal
End Sub
Private Sub txtInscrMunicipal_LostFocus()
  Pintar_Controle txtInscrMunicipal, tpCorContr_Normal
End Sub

Private Sub txtTelefone_GotFocus()
  Seleciona_Conteudo_Controle txtTelefone
End Sub
Private Sub txtTelefone_LostFocus()
  Pintar_Controle txtTelefone, tpCorContr_Normal
End Sub

Private Sub txtTitulo_GotFocus()
  Seleciona_Conteudo_Controle txtTitulo
End Sub
Private Sub txtTitulo_LostFocus()
  Pintar_Controle txtTitulo, tpCorContr_Normal
End Sub

