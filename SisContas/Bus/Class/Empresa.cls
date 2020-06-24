VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirEmpresa(strNome As String, _
                          strTel As String, _
                          strTel1 As String, _
                          strTel2 As String, _
                          strContato As String, _
                          strObs As String, _
                          strTipoEmpresaId As String, _
                          strPercentualAg As String, _
                          strEndRua As String, _
                          strEndNumero As String, _
                          strEndComplemento As String, _
                          strEndCep As String, _
                          strEndBairro As String, _
                          strEndCidade As String, _
                          strEndEstado As String, _
                          strEndPais As String, _
                          strCobRua As String, _
                          strCobNumero As String, _
                          strCobComplemento As String, _
                          strCobCep As String, _
                          strCobBairro As String, _
                          strCobCidade As String, _
                          strCobEstado As String, strCobPais As String, _
                          strCGC As String, strInscrEstadual As String, strInscrMunicipal As String, strPercDescDiaria As String, strCredito As String, _
                          strNomeFantasia As String, strNroContrato As String, strDtInicioContrato As String, strDtFimContrato As String, blnEndCobIgualEndCorr As Boolean, lngParceiroId As Long)

  On Error GoTo trata
  Dim objEmpresa As datSisContas.clsEmpresa
  '
  Set objEmpresa = New datSisContas.clsEmpresa
  
  'Inserir Empresa
  objEmpresa.InserirEmpresa strNome, _
                            strTel, _
                            strTel1, _
                            strTel2, _
                            strContato, _
                            strObs, _
                            strTipoEmpresaId, _
                            strPercentualAg, _
                            strEndRua, _
                            strEndNumero, _
                            strEndComplemento, _
                            strEndCep, _
                            strEndBairro, _
                            strEndCidade, _
                            strEndEstado, _
                            strEndPais, _
                            strCobRua, _
                            strCobNumero, _
                            strCobComplemento, _
                            strCobCep, _
                            strCobBairro, _
                            strCobCidade, _
                            strCobEstado, strCobPais, _
                            strCGC, strInscrEstadual, strInscrMunicipal, strPercDescDiaria, strCredito, _
                            strNomeFantasia, strNroContrato, strDtInicioContrato, strDtFimContrato, blnEndCobIgualEndCorr, lngParceiroId
  '
  Set objEmpresa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarEmpresa(lngEMPRESAID As Long, strNome As String, _
                          strTel As String, _
                          strTel1 As String, _
                          strTel2 As String, _
                          strContato As String, _
                          strObs As String, _
                          strTipoEmpresaId As String, _
                          strPercentualAg As String, _
                          strEndRua As String, _
                          strEndNumero As String, _
                          strEndComplemento As String, _
                          strEndCep As String, _
                          strEndBairro As String, _
                          strEndCidade As String, _
                          strEndEstado As String, _
                          strEndPais As String, _
                          strCobRua As String, _
                          strCobNumero As String, _
                          strCobComplemento As String, _
                          strCobCep As String, _
                          strCobBairro As String, _
                          strCobCidade As String, strCobEstado As String, strCobPais As String, _
                          strCGC As String, strInscrEstadual As String, strInscrMunicipal As String, strPercDescDiaria As String, strCredito As String, _
                          strNomeFantasia As String, strNroContrato As String, strDtInicioContrato As String, strDtFimContrato As String, blnEndCobIgualEndCorr As Boolean)
                         
  On Error GoTo trata
  Dim objEmpresa As datSisContas.clsEmpresa
  Set objEmpresa = New datSisContas.clsEmpresa
  
  'Inserir Empresa
  objEmpresa.AlterarEmpresa lngEMPRESAID, strNome, _
                            strTel, _
                            strTel1, _
                            strTel2, _
                            strContato, _
                            strObs, _
                            strTipoEmpresaId, _
                            strPercentualAg, _
                            strEndRua, _
                            strEndNumero, _
                            strEndComplemento, _
                            strEndCep, _
                            strEndBairro, _
                            strEndCidade, _
                            strEndEstado, _
                            strEndPais, _
                            strCobRua, _
                            strCobNumero, _
                            strCobComplemento, _
                            strCobCep, _
                            strCobBairro, _
                            strCobCidade, _
                            strCobEstado, strCobPais, _
                            strCGC, strInscrEstadual, strInscrMunicipal, strPercDescDiaria, strCredito, _
                            strNomeFantasia, strNroContrato, strDtInicioContrato, strDtFimContrato, blnEndCobIgualEndCorr

  '
  Set objEmpresa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarEmpresa(lngEMPRESAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objEmpresa As datSisContas.clsEmpresa
  Set objEmpresa = New datSisContas.clsEmpresa
  
  'Listar Empresa
  Set ListarEmpresa = objEmpresa.ListarEmpresa(lngEMPRESAID)
  '
  Set objEmpresa = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function ListarEmpresaPeloNome(strNome As String, _
                                      lngEMPRESAID As Long, _
                                      lngParceiroId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objEmpresa As datSisContas.clsEmpresa
  Set objEmpresa = New datSisContas.clsEmpresa
  
  'Listar Empresa pelo nome
  Set ListarEmpresaPeloNome = objEmpresa.ListarEmpresaPeloNome(strNome, _
                                                               lngEMPRESAID, _
                                                               lngParceiroId)
  '
  Set objEmpresa = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirEmpresa(lngEMPRESAID As Long)

  On Error GoTo trata
  Dim objEmpresa As datSisContas.clsEmpresa
  Set objEmpresa = New datSisContas.clsEmpresa
  
  'Excluir Empresa
  objEmpresa.ExlcuirEmpresa lngEMPRESAID
  '
  Set objEmpresa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

