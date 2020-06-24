VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFichaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarFichaClientePeloPkid(lngFichaClienteId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objFichaCliente As datSisContas.clsFichaCliente
  Set objFichaCliente = New datSisContas.clsFichaCliente
  
  'Listar Ficha Cliente
  Set SelecionarFichaClientePeloPkid = objFichaCliente.SelecionarFichaClientePeloPkid(lngFichaClienteId)
  '
  Set objFichaCliente = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarFichaCliente(lngTIPODOCUMENTOID As Long, _
                                       strNroDocumento As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objFichaCliente As datSisContas.clsFichaCliente
  Set objFichaCliente = New datSisContas.clsFichaCliente
  
  'Listar Ficha Cliente
  Set SelecionarFichaCliente = objFichaCliente.SelecionarFichaCliente(lngTIPODOCUMENTOID, _
                                                                      strNroDocumento)
  '
  Set objFichaCliente = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarFichaClientePeloSobreNome(strSobreNome As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objFichaCliente As datSisContas.clsFichaCliente
  Set objFichaCliente = New datSisContas.clsFichaCliente
  
  'Listar Ficha Cliente
  Set SelecionarFichaClientePeloSobreNome = objFichaCliente.SelecionarFichaClientePeloSobreNome(strSobreNome)
  '
  Set objFichaCliente = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirFichaCliente(lngFichaClienteId As Long)
  On Error GoTo trata
  Dim objFichaCliente As datSisContas.clsFichaCliente
  Set objFichaCliente = New datSisContas.clsFichaCliente
  
  'Excluir Ficha Cliente
  objFichaCliente.ExcluirFichaCliente lngFichaClienteId
  '
  Set objFichaCliente = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Function InserirFichaCliente(lngTIPODOCUMENTOID As Long, _
                                    strSobreNome As String, _
                                    strNome As String, _
                                    strEndereco As String, _
                                    strNumero As String, _
                                    strComplemento As String, _
                                    strBairro As String, _
                                    strCidade As String, _
                                    strEstado As String, _
                                    strCep As String, _
                                    strPais As String, _
                                    strTel1 As String, _
                                    strTel2 As String, _
                                    strTel3 As String, _
                                    strDtNascimento As String, _
                                    strSexo As String, _
                                    strNroDocumento As String, _
                                    strPercDesc As String, _
                                    strEmail As String, _
                                    strObservacao As String) As Long

  On Error GoTo trata
  Dim objFichaCliente     As datSisContas.clsFichaCliente
  Dim lngFichaClinteId    As Long
  '
  Set objFichaCliente = New datSisContas.clsFichaCliente
  'Inserir Ficha Cliente
  lngFichaClinteId = objFichaCliente.InserirFichaCliente(lngTIPODOCUMENTOID, _
                                                         strSobreNome, _
                                                         strNome, _
                                                         strEndereco, _
                                                         strNumero, _
                                                         strComplemento, _
                                                         strBairro, _
                                                         strCidade, _
                                                         strEstado, _
                                                         strCep, _
                                                         strPais, _
                                                         strTel1, _
                                                         strTel2, _
                                                         strTel3, _
                                                         strDtNascimento, _
                                                         strSexo, _
                                                         strNroDocumento, _
                                                         strPercDesc, _
                                                         strEmail, _
                                                         strObservacao)
  '
  Set objFichaCliente = Nothing
  InserirFichaCliente = lngFichaClinteId
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function CadastrarFichaCliente(lngTIPODOCUMENTOID As Long, _
                                      strSobreNome As String, _
                                      strNome As String, _
                                      strEndereco As String, _
                                      strNumero As String, _
                                      strComplemento As String, _
                                      strBairro As String, _
                                      strCidade As String, _
                                      strEstado As String, _
                                      strCep As String, _
                                      strPais As String, _
                                      strTel1 As String, _
                                      strTel2 As String, _
                                      strTel3 As String, _
                                      strDtNascimento As String, _
                                      strSexo As String, _
                                      strNroDocumento As String, _
                                      strPercDesc As String, _
                                      strEmail As String, _
                                      strObservacao As String) As Long

  On Error GoTo trata
  Dim objFichaCliente     As datSisContas.clsFichaCliente
  Dim lngFichaClinteId    As Long
  Dim objRs               As ADODB.Recordset
  '
  lngFichaClinteId = 0
  CadastrarFichaCliente = 0
  Set objFichaCliente = New datSisContas.clsFichaCliente
  Set objRs = objFichaCliente.SelecionarFichaCliente(lngTIPODOCUMENTOID, _
                                                     strNroDocumento)
  If Not objRs.EOF Then
    If IsNumeric(objRs.Fields("PKID").Value) Then
      lngFichaClinteId = objRs.Fields("PKID").Value
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  If lngFichaClinteId = 0 Then
    'Inserir Ficha Cliente
    lngFichaClinteId = objFichaCliente.InserirFichaCliente(lngTIPODOCUMENTOID, _
                                                           strSobreNome, _
                                                           strNome, _
                                                           strEndereco, _
                                                           strNumero, _
                                                           strComplemento, _
                                                           strBairro, _
                                                           strCidade, _
                                                           strEstado, _
                                                           strCep, _
                                                           strPais, _
                                                           strTel1, _
                                                           strTel2, _
                                                           strTel3, _
                                                           strDtNascimento, _
                                                           strSexo, _
                                                           strNroDocumento, _
                                                           strPercDesc, _
                                                           strEmail, _
                                                           strObservacao)
  Else
    'Atualizar Ficha Cliente
    objFichaCliente.AlterarFichaCliente lngFichaClinteId, _
                                        lngTIPODOCUMENTOID, _
                                        strSobreNome, _
                                        strNome, _
                                        strEndereco, _
                                        strNumero, _
                                        strComplemento, _
                                        strBairro, _
                                        strCidade, _
                                        strEstado, _
                                        strCep, _
                                        strPais, _
                                        strTel1, _
                                        strTel2, _
                                        strTel3, _
                                        strDtNascimento, _
                                        strSexo, _
                                        strNroDocumento, _
                                        strPercDesc, _
                                        strEmail, _
                                        strObservacao
  End If
  '
  Set objFichaCliente = Nothing
  CadastrarFichaCliente = lngFichaClinteId
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

