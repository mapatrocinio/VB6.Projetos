VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProfissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ListarProfissaoPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objProfissao As datApler.clsProfissao
  Set objProfissao = New datApler.clsProfissao
  
  'Listar Profissao pela descrição
  Set ListarProfissaoPelaDesc = objProfissao.ListarProfissaoPelaDesc(strDescricao)
  '
  Set objProfissao = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub InserirProfissao(strDescricao As String)

  On Error GoTo trata
  Dim objProfissao As datApler.clsProfissao
  '
  Set objProfissao = New datApler.clsProfissao
  
  'Inserir Profissao
  objProfissao.InserirProfissao strDescricao
  '
  Set objProfissao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarProfissao(lngPROFISSAOID As Long, _
                            strDescricao As String)
                         
  On Error GoTo trata
  Dim objProfissao As datApler.clsProfissao
  Set objProfissao = New datApler.clsProfissao
  
  'Inserir Profissao
  objProfissao.AlterarProfissao lngPROFISSAOID, _
                                strDescricao

  '
  Set objProfissao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarProfissao(lngPROFISSAOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objProfissao As datApler.clsProfissao
  Set objProfissao = New datApler.clsProfissao
  
  'Listar Profissao
  Set ListarProfissao = objProfissao.ListarProfissao(lngPROFISSAOID)
  '
  Set objProfissao = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirProfissao(lngPROFISSAOID As Long)

  On Error GoTo trata
  Dim objProfissao As datApler.clsProfissao
  Set objProfissao = New datApler.clsProfissao
  
  'Excluir Profissao
  objProfissao.ExlcuirProfissao lngPROFISSAOID
  '
  Set objProfissao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

