VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConfiguracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirConfiguracao(ByVal lngCONFIGURACAOID As Long)
  Dim objConf As datSisMed.clsConfiguracao
  '
  On Error GoTo trata
  Set objConf = New datSisMed.clsConfiguracao
  '
  objConf.ExcluirConfiguracao lngCONFIGURACAOID
  '
  Set objConf = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function VerificaExclusaoConfiguracao(lngCONFIGURACAOID As Long, _
                                             ByRef strDescricaoErro As String) As Boolean
  On Error GoTo trata
  Dim objConfig As datSisMed.clsConfiguracao
  Set objConfig = New datSisMed.clsConfiguracao

  'Verificar
  VerificaExclusaoConfiguracao = objConfig.VerificaExclusaoConfiguracao(lngCONFIGURACAOID, _
                                                                        strDescricaoErro)
  '
  Set objConfig = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub InserirConfiguracao()
                         

  Dim objConfiguracao As datSisMed.clsConfiguracao
  Set objConfiguracao = New datSisMed.clsConfiguracao
  
  'Inserir Configuracao
  objConfiguracao.InserirConfiguracao
  '
  Set objConfiguracao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarConfiguracaoCadastro(lngCONFIGURACAOID As Long, _
                                       strEmpresa As String, _
                                       strTitulo As String, _
                                       strCnpj As String, _
                                       strInscrMunicipal As String, _
                                       strEndereco As String, _
                                       strBairro As String, _
                                       strCidade As String, _
                                       strEstado As String, _
                                       strCep As String, _
                                       strTel As String, _
                                       strPathLocal As String, _
                                       strPathLocalBackup As String, _
                                       strPathRede As String, _
                                       strMaxDiasAtend As String, _
                                       blnTrabComImpA5 As Boolean)
                         

  Dim blnConfiguracao As datSisMed.clsConfiguracao
  Set blnConfiguracao = New datSisMed.clsConfiguracao
  
  'Inserir Configuracao - Cadastro
  blnConfiguracao.AlterarConfiguracaoCadastro lngCONFIGURACAOID, _
                                              strEmpresa, _
                                              strTitulo, _
                                              strCnpj, _
                                              strInscrMunicipal, _
                                              strEndereco, _
                                              strBairro, _
                                              strCidade, _
                                              strEstado, _
                                              strCep, _
                                              strTel, _
                                              strPathLocal, _
                                              strPathLocalBackup, _
                                              strPathRede, _
                                              strMaxDiasAtend, _
                                              blnTrabComImpA5

  '
  Set blnConfiguracao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarConfiguracaoCadastro(lngCONFIGURACAOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objConfiguracao As datSisMed.clsConfiguracao
  Set objConfiguracao = New datSisMed.clsConfiguracao
  
  'Listar Configuracao - Cadastro
  Set ListarConfiguracaoCadastro = objConfiguracao.ListarConfiguracaoCadastro(lngCONFIGURACAOID)
  '
  Set objConfiguracao = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ListarConfiguracao() As ADODB.Recordset
  On Error GoTo trata
  Dim objConfiguracao As datSisMed.clsConfiguracao
  Set objConfiguracao = New datSisMed.clsConfiguracao
  
  'Listar Configuracao
  Set ListarConfiguracao = objConfiguracao.ListarConfiguracao
  '
  Set objConfiguracao = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

