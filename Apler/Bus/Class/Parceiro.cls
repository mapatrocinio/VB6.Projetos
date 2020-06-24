VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParceiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarParceiroPeloNome(strNome As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objParceiro As datApler.clsParceiro
  Set objParceiro = New datApler.clsParceiro
  
  'Listar Parceiro
  Set SelecionarParceiroPeloNome = objParceiro.SelecionarParceiroPeloNome(strNome)
  '
  Set objParceiro = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarParceiroPeloPkid(lngPARCEIROID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objParceiro As datApler.clsParceiro
  Set objParceiro = New datApler.clsParceiro
  
  'Listar Parceiro
  Set SelecionarParceiroPeloPkid = objParceiro.SelecionarParceiroPeloPkid(lngPARCEIROID)
  '
  Set objParceiro = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirParceiro(lngPARCEIROID As Long)
  On Error GoTo trata
  Dim objParceiro As datApler.clsParceiro
  Set objParceiro = New datApler.clsParceiro
  
  'Excluir Parceiro
  objParceiro.ExcluirParceiro lngPARCEIROID
  '
  Set objParceiro = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirParceiro(ByRef lngPKID As Long, _
                           lngCONVENIOID As Long, _
                           strNome As String, _
                           strNomeFantasia As String, _
                           strCnpj As String, _
                           strInscrEstd As String, _
                           strInscrMunic As String, _
                           strTel1 As String, strTel2 As String, strTel3 As String, strFax As String, _
                           strEmail As String, _
                           strContato As String, _
                           strTelContato As String, _
                           strEndRua As String, _
                           strEndNumero As String, _
                           strEndCompl As String, _
                           strEndEstado As String, _
                           strEndCep As String, _
                           strEndBairro As String, _
                           strEndCidade As String, _
                           strStatus As String, _
                           strStatusAnterior As String)
  On Error GoTo trata
  Dim objParceiro     As datApler.clsParceiro
  '
  Set objParceiro = New datApler.clsParceiro
  'Inserir Parceiro
  objParceiro.InserirParceiro lngPKID, _
                              lngCONVENIOID, _
                              strNome, _
                              strNomeFantasia, _
                              strCnpj, _
                              strInscrEstd, _
                              strInscrMunic, _
                              strTel1, strTel2, strTel3, strFax, _
                              strEmail, _
                              strContato, _
                              strTelContato, _
                              strEndRua, _
                              strEndNumero, _
                              strEndCompl, _
                              strEndEstado, _
                              strEndCep, _
                              strEndBairro, _
                              strEndCidade, _
                              strStatus, _
                              strStatusAnterior
  '
  Set objParceiro = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarParceiro(lngPKID As Long, _
                           lngCONVENIOID As Long, _
                           strNome As String, _
                           strNomeFantasia As String, _
                           strCnpj As String, _
                           strInscrEstd As String, _
                           strInscrMunic As String, _
                           strTel1 As String, strTel2 As String, strTel3 As String, strFax As String, _
                           strEmail As String, _
                           strContato As String, _
                           strTelContato As String, _
                           strEndRua As String, _
                           strEndNumero As String, _
                           strEndCompl As String, _
                           strEndEstado As String, _
                           strEndCep As String, _
                           strEndBairro As String, _
                           strEndCidade As String, _
                           strStatus As String, _
                           strStatusAnterior As String)

  On Error GoTo trata
  Dim objParceiro     As datApler.clsParceiro
  '
  Set objParceiro = New datApler.clsParceiro
  'Atualizar Parceiro
  objParceiro.AlterarParceiro lngPKID, _
                              lngCONVENIOID, _
                              strNome, _
                              strNomeFantasia, _
                              strCnpj, _
                              strInscrEstd, _
                              strInscrMunic, _
                              strTel1, strTel2, strTel3, strFax, _
                              strEmail, _
                              strContato, _
                              strTelContato, _
                              strEndRua, _
                              strEndNumero, _
                              strEndCompl, _
                              strEndEstado, _
                              strEndCep, _
                              strEndBairro, _
                              strEndCidade, _
                              strStatus, _
                              strStatusAnterior
  '
  Set objParceiro = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

