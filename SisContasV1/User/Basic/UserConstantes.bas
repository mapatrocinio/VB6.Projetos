Attribute VB_Name = "mdlUserConstantes"
Option Explicit

Public Const TITULOSISTEMA = "Sistema Gerenciador de Contas"

Public gsBDadosPath   As String
Public gsReportPath   As String
Public gsBMPPath      As String
Public gsIconsPath    As String
Public gsAppPath      As String
Public gsBMP          As String
'Definido no form load
Public ConnectRpt     As String
                                          
Enum tpStatus
  tpStatus_Incluir
  tpStatus_Alterar
  tpStatus_Consultar
End Enum

Enum tpCorControle
  tpCorContr_Erro = &HC0FFFF
  tpCorContr_Normal = vbWhite
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
End Enum
'-------MODO ANTIGO
'Utilizado para guardar nome do usuário que liberou ação
Global gsNomeUsuLib As String
Global gsNivelUsuLib As String

Global gsNomeEmpresa As String
Global gsNomeUsu As String
Global gsNivel As String
Global gsPathBackup As String
'Níveis de acesso
Global Const gsDiretor = "DIR"
Global Const gsGerente = "GER"
Global Const gsRecepcao = "REC"
Global Const gsPortaria = "POR"
Global Const gsAdmin = "ADM"
Global Const gsEstoque = "EST"

Public Const nomeBDados = "SisMotel.MDB"

Enum TpObriga
  TpObrigatorio
  TpNaoObrigatorio
End Enum

Enum TpTipoDados
  tpDados_Texto '0 a 255
  tpDados_Memo 'Sem Limite
  tpDados_Inteiro '-32767 a 32767
  tpDados_Longo 'Sem limite
  tpDados_DataHora 'MM/DD/YYYY hh:mm:ss
  tpDados_Moeda '121212.98
  tpDados_Boolean '121212.98
End Enum

Enum tpAceitaNulo
  tpNulo_Aceita
  tpNulo_NaoAceita
End Enum

