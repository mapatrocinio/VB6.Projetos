VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQualificacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ListarQualificacaoPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objQualificacao As datApler.clsQualificacao
  Set objQualificacao = New datApler.clsQualificacao
  
  'Listar Qualificacao pela descrição
  Set ListarQualificacaoPelaDesc = objQualificacao.ListarQualificacaoPelaDesc(strDescricao)
  '
  Set objQualificacao = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub InserirQualificacao(strDescricao As String)

  On Error GoTo trata
  Dim objQualificacao As datApler.clsQualificacao
  '
  Set objQualificacao = New datApler.clsQualificacao
  
  'Inserir Qualificacao
  objQualificacao.InserirQualificacao strDescricao
  '
  Set objQualificacao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarQualificacao(lngQUALIFICACAOID As Long, _
                         strDescricao As String)
                         
  On Error GoTo trata
  Dim objQualificacao As datApler.clsQualificacao
  Set objQualificacao = New datApler.clsQualificacao
  
  'Inserir Qualificacao
  objQualificacao.AlterarQualificacao lngQUALIFICACAOID, _
                          strDescricao

  '
  Set objQualificacao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarQualificacao(lngQUALIFICACAOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objQualificacao As datApler.clsQualificacao
  Set objQualificacao = New datApler.clsQualificacao
  
  'Listar Qualificacao
  Set ListarQualificacao = objQualificacao.ListarQualificacao(lngQUALIFICACAOID)
  '
  Set objQualificacao = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirQualificacao(lngQUALIFICACAOID As Long)

  On Error GoTo trata
  Dim objQualificacao As datApler.clsQualificacao
  Set objQualificacao = New datApler.clsQualificacao
  
  'Excluir Qualificacao
  objQualificacao.ExlcuirQualificacao lngQUALIFICACAOID
  '
  Set objQualificacao = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

