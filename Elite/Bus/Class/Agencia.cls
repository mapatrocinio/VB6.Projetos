VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAgencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarAgenciaPeloNome(strNome As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objAgencia As datElite.clsAgencia
  Set objAgencia = New datElite.clsAgencia
  
  'Listar Agencia
  Set SelecionarAgenciaPeloNome = objAgencia.SelecionarAgenciaPeloNome(strNome)
  '
  Set objAgencia = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarAgenciaPeloPkid(lngAGENCIAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objAgencia As datElite.clsAgencia
  Set objAgencia = New datElite.clsAgencia
  
  'Listar Agencia
  Set SelecionarAgenciaPeloPkid = objAgencia.SelecionarAgenciaPeloPkid(lngAGENCIAID)
  '
  Set objAgencia = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirAgencia(lngAGENCIAID As Long)
  On Error GoTo trata
  Dim objAgencia As datElite.clsAgencia
  Set objAgencia = New datElite.clsAgencia
  
  'Excluir Agencia
  objAgencia.ExcluirAgencia lngAGENCIAID
  '
  Set objAgencia = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirAgencia(lngPKID As Long, _
                          strNome As String, _
                          strTel1 As String, strTel2 As String, strTel3 As String, _
                          strEndRua As String, _
                          strEndNumero As String, _
                          strEndCompl As String, _
                          strEndEstado As String, _
                          strEndCep As String, _
                          strEndBairro As String, _
                          strEndCidade As String, _
                          strObservacao As String, _
                          strStatus As String)
  On Error GoTo trata
  Dim objAgencia     As datElite.clsAgencia
  '
  Set objAgencia = New datElite.clsAgencia
  'Inserir Agencia
  objAgencia.InserirAgencia lngPKID, _
                            strNome, _
                            strTel1, strTel2, strTel3, _
                            strEndRua, _
                            strEndNumero, _
                            strEndCompl, _
                            strEndEstado, _
                            strEndCep, _
                            strEndBairro, _
                            strEndCidade, _
                            strObservacao, _
                            strStatus
  '
  Set objAgencia = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarAgencia(lngPKID As Long, _
                          strNome As String, _
                          strTel1 As String, strTel2 As String, strTel3 As String, _
                          strEndRua As String, _
                          strEndNumero As String, _
                          strEndCompl As String, _
                          strEndEstado As String, _
                          strEndCep As String, _
                          strEndBairro As String, _
                          strEndCidade As String, _
                          strObservacao As String, _
                          strStatus As String)

  On Error GoTo trata
  Dim objAgencia     As datElite.clsAgencia
  '
  Set objAgencia = New datElite.clsAgencia
  'Atualizar Agencia
  objAgencia.AlterarAgencia lngPKID, _
                            strNome, _
                            strTel1, strTel2, strTel3, _
                            strEndRua, _
                            strEndNumero, _
                            strEndCompl, _
                            strEndEstado, _
                            strEndCep, _
                            strEndBairro, _
                            strEndCidade, _
                            strObservacao, _
                            strStatus
  '
  Set objAgencia = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

