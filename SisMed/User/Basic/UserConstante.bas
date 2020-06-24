Attribute VB_Name = "mdlUserConstante"
Option Explicit

'Variáveis de configur~ção
Public Const gbTrabComTroco As Boolean = True
Public Const gbTrabComChequesBons As Boolean = False
Public Const intQtdPontos As Integer = 10
Public Const gbTrabComDepSangria As Boolean = False
''''Public Const gsNomeEmpresa As String = "Casa de Saúde São Francisco de Paula"
Global glMaxSequencialImp As Long
'Variáveis utilizadas para o scaner
Public gsPathLocal As String
Public gsPathLocalBackup As String
Public gsPathRede As String
Public giMaxDiasAtend As Integer

Global gbTrabImpA5 As Boolean


Public gsBDadosPath   As String
Public gsReportPath   As String
Public gsBMPPath      As String
Public gsIconsPath    As String
Public gsAppPath      As String
Public gsBMP          As String
Public ConnectRpt     As String

Global gsNomeServidorBD As String
Global lngCONFIGURACAOID As Long

Public Const nomeBDados = "SisMed.MDB"
Public Const psUsuarioPerm = "Admin"
Public Const psSenhaPerm = "SHOGUM2806"
Public Const psSenhaPermWks = "" '"BKDR15864"

Global giFuncionarioId As Long
Global gbTrabComScaner As Boolean


'Utilizado para guardar nome do usuário que liberou ação
Global gsNomeUsuLib As String
Global gsNivelUsuLib As String

Global gsNomeEmpresa As String
Global gsNomeUsu As String
Global gsNomeUsuCompleto As String
Global gsNivel As String
Global gsPathBackup As String

'Níveis de acesso
Global Const gsAdmin = "ADM"
Global Const gsDiretor = "DIR"
Global Const gsGerente = "GER"
Global Const gsCaixa = "CAI"
Global Const gsLaboratorio = "LAB"
Global Const gsFinanceiro = "FIN"
Global Const gsArquivista = "ARQ"
Global Const gsPrestador = "PRE"


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

Enum tpIcProntuario
  tpIcProntuario_Func = 0
  tpIcProntuario_Pac = 1
  tpIcProntuario_Prest = 2
End Enum


