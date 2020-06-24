VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCartao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirCartao(strNome As String, _
                         strNumero As String, _
                         strPercTxAdm As String)

  On Error GoTo trata
  Dim objCartao As datSisMetal.clsCartao
  '
  Set objCartao = New datSisMetal.clsCartao
  
  'Inserir Cartao
  objCartao.InserirCartao strNome, _
                          strNumero, _
                          strPercTxAdm
  '
  Set objCartao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarCartao(lngCARTAOID As Long, _
                         strNome As String, _
                         strNumero As String, _
                         strPercTxAdm As String)
                         
  On Error GoTo trata
  Dim objCartao As datSisMetal.clsCartao
  Set objCartao = New datSisMetal.clsCartao
  
  'Inserir Cartao
  objCartao.AlterarCartao lngCARTAOID, _
                          strNome, _
                          strNumero, _
                          strPercTxAdm

  '
  Set objCartao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarCartao(lngCARTAOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objCartao As datSisMetal.clsCartao
  Set objCartao = New datSisMetal.clsCartao
  
  'Listar Cartao
  Set ListarCartao = objCartao.ListarCartao(lngCARTAOID)
  '
  Set objCartao = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirCartao(lngCARTAOID As Long)

  On Error GoTo trata
  Dim objCartao As datSisMetal.clsCartao
  Set objCartao = New datSisMetal.clsCartao
  
  'Excluir Cartao
  objCartao.ExlcuirCartao lngCARTAOID
  '
  Set objCartao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

