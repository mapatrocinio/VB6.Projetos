Attribute VB_Name = "mdlUserConstante"
Option Explicit

Public Const TITULOSISTEMA = "Sistema Gerenciador de Conv�nios"


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

'Utilizado para guardar nome do usu�rio que liberou a��o
Global gsNomeUsuLib As String
Global gsNivelUsuLib As String

Global gsNomeEmpresa As String
Global gsNomeUsu As String
Global gsNivel As String
Global gsPathBackup As String

'N�veis de acesso
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

