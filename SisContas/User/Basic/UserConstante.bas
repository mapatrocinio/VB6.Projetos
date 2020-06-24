Attribute VB_Name = "mdlUserConstante"
Option Explicit

'Importadas mas sem configuração
Global blnClienteTrabMultiplasConfig As Boolean
Global gbPedirSenhaSupLibChqReceb As Boolean
Global gbTrabComChequesBons As Boolean
Global giQtdDiasParaCompensar As Integer
Global giQtdChequesBons As Integer
Global glMaxSequencialImp As Long
Global lngCONFIGURACAOID As Long
Global glDiaFechaFolha As Integer
Global gbTrabComTroco As Boolean
Global gbTrabComLiberacao As Boolean
Global gbTrabSaida As Boolean
Global gbTrabSuiteAptoLimpo As Boolean
Global gbTrabComImpFiscal As Boolean

Global glParceiroId As Long


Global giTpTipo As TpTipo
'constantes para modos de edição
Public Const MODONORMAL As Byte = 0
Public Const MODOALTERAR As Byte = 1
Public Const MODOINSERIR As Byte = 2
Public Const MODOEXCLUIR As Byte = 3

Public gsBDadosPath   As String
Public gsReportPath   As String
Public gsBMPPath      As String
Public gsIconsPath    As String
Public gsAppPath      As String
Public gsBMP          As String
Public ConnectRpt     As String

Public Const nomeBDados = "Apler.MDB"
Public Const psUsuarioPerm = "Admin"
Public Const psSenhaPerm = "SHOGUM2806"
Public Const psSenhaPermWks = "" '"BKDR15864"

'Utilizado para guardar nome do usuário que liberou ação
Global gsNomeUsuLib As String
Global gsNivelUsuLib As String

Global gsNomeEmpresa As String
Global gsNomeUsu As String
Global gsNivel As String
Global gsPathBackup As String
Global gsNomeServidorBD As String

'Níveis de acesso
Global Const gsDiretor = "DIR"
Global Const gsGerente = "GER"
Global Const gsRecepcao = "REC"
Global Const gsPortaria = "POR"
Global Const gsAdmin = "ADM"
Global Const gsEstoque = "EST"
Global Const gsFinanceiro = "FIN"

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

Enum tpMaskValor
  TpMaskData
  TpMaskMoeda
  TpMaskLongo
  TpMaskOutros
  TpMaskSemMascara
End Enum
Enum TpTipo
  TpTipo_Motel = 0
  TpTipo_Pousada = 1
  TpTipo_Hotel = 2
End Enum

