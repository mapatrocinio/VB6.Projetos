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

Public Function SelecionarEmpresaPeloNome(strNome As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objEmpresa As datSisLoc.clsEmpresa
  Set objEmpresa = New datSisLoc.clsEmpresa
  
  'Listar Empresa
  Set SelecionarEmpresaPeloNome = objEmpresa.SelecionarEmpresaPeloNome(strNome)
  '
  Set objEmpresa = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarEmpresaPeloPkid(lngEMPRESAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objEmpresa As datSisLoc.clsEmpresa
  Set objEmpresa = New datSisLoc.clsEmpresa
  
  'Listar Empresa
  Set SelecionarEmpresaPeloPkid = objEmpresa.SelecionarEmpresaPeloPkid(lngEMPRESAID)
  '
  Set objEmpresa = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirEmpresa(lngEMPRESAID As Long)
  On Error GoTo trata
  Dim objEmpresa As datSisLoc.clsEmpresa
  Set objEmpresa = New datSisLoc.clsEmpresa
  
  'Excluir Empresa
  objEmpresa.ExcluirEmpresa lngEMPRESAID
  '
  Set objEmpresa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirEmpresa(ByRef lngPKID As Long, _
                           lngTIPOEMPRESAID As Long, _
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
                           strStatus As String)
  On Error GoTo trata
  Dim objEmpresa     As datSisLoc.clsEmpresa
  '
  Set objEmpresa = New datSisLoc.clsEmpresa
  'Inserir Empresa
  objEmpresa.InserirEmpresa lngPKID, _
                              lngTIPOEMPRESAID, _
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
                              strStatus
  '
  Set objEmpresa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarEmpresa(lngPKID As Long, _
                           lngTIPOEMPRESAID As Long, _
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
                           strStatus As String)

  On Error GoTo trata
  Dim objEmpresa     As datSisLoc.clsEmpresa
  '
  Set objEmpresa = New datSisLoc.clsEmpresa
  'Atualizar Empresa
  objEmpresa.AlterarEmpresa lngPKID, _
                              lngTIPOEMPRESAID, _
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
                              strStatus
  '
  Set objEmpresa = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

