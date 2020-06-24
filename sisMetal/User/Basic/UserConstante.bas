Attribute VB_Name = "mdlUserConstante"
Option Explicit

Global gsCaminhoImagemCompra    As String
Global giQtdDiasVenda           As Integer
Global giQtdDiasPedido          As Integer

Public gsBDadosPath   As String
Public gsReportPath   As String
Public gsBMPPath      As String
Public gsIconsPath    As String
Public gsAppPath      As String
Public gsBMP          As String
Public ConnectRpt     As String

Global gsNomeServidorBD As String
Global giFuncionarioId As Long
Global gsNomeUsuCompleto As String


Public Const nomeBDados = "SisMetal.MDB"
Public Const psUsuarioPerm = "Admin"
Public Const psSenhaPerm = "SHOGUM2806"
Public Const psSenhaPermWks = "" '"BKDR15864"

'Utilizado para guardar nome do usuário que liberou ação
Global gsNomeUsuLib As String
Global gsNivelUsuLib As String
Global giFunIdUsuLib As Long

Global gsNomeEmpresa As String
Global gsNomeUsu As String
Global gsNivel As String
Global gsPathBackup As String

'Níveis de acesso
Global Const gsAdmin = "ADM"
Global Const gsDiretor = "DIR"
Global Const gsGerente = "GER"
Global Const gsCompra = "COM"
Global Const gsCaixa = "CAI"
Global Const gsFinanceiro = "FIN"
Global Const gsLoja = "LOJ"
Global Const gsSemAcesso = "SEM"
Global Const gsVendedor = "VEN"


Enum TpObriga
  TpObrigatorio
  TpNaoObrigatorio
End Enum

Enum tpStatus
  tpStatus_Incluir
  tpStatus_Alterar
  tpStatus_Consultar
End Enum

Enum tpCorControle
  tpCorContr_Erro = &HC0FFFF
  tpCorContr_Normal = vbWhite
  tpCorContr_Desabilitado = &HE0E0E0
End Enum
Enum tpBmpForm
  tpBmp_Vazio = "0"
  tpBmp_Login = "1"
  tpBmp_MDI = "2"
End Enum

Enum tpIcPessoa
  tpIcPessoa_Func = 0
  tpIcPessoa_Pac = 1
  tpIcPessoa_Prest = 2
End Enum

Enum tpTipoVenda
  tpTipoVenda_Balc = 0
  tpTipoVenda_Clie = 1
  tpTipoVenda_Emp = 2
End Enum


Enum tpMaskValor
  TpMaskData
  TpMaskMoeda
  TpMaskLongo
  TpMaskOutros
  TpMaskSemMascara
End Enum


