VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrupoProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarGrupoProduto(lngGRUPO_PRODUTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objGrupoProduto As datSisMetal.clsGrupoProduto
  Set objGrupoProduto = New datSisMetal.clsGrupoProduto
  
  'Listar GrupoProduto
  Set SelecionarGrupoProduto = objGrupoProduto.SelecionarGrupoProduto(lngGRUPO_PRODUTOID)
  '
  Set objGrupoProduto = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Sub IncluirGrupoProduto(strNome As String, _
                        lngREFERENCIAPRODUTOID As Long)
  
  On Error GoTo trata
  Dim objGrupoProduto As datSisMetal.clsGrupoProduto
  Set objGrupoProduto = New datSisMetal.clsGrupoProduto
  
  'Inserir GrupoProduto
  objGrupoProduto.IncluirGrupoProduto strNome, _
                        lngREFERENCIAPRODUTOID
  '
  Set objGrupoProduto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarGrupoProduto(strNome As String, _
                        lngREFERENCIAPRODUTOID As Long, _
                        lngGRUPO_PRODUTOID As Long)
  On Error GoTo trata
  Dim objGrupoProduto As datSisMetal.clsGrupoProduto
  Set objGrupoProduto = New datSisMetal.clsGrupoProduto
  
  'Alterar GrupoProduto
  objGrupoProduto.AlterarGrupoProduto strNome, _
                        lngREFERENCIAPRODUTOID, _
                        lngGRUPO_PRODUTOID

  '
  Set objGrupoProduto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub
Public Sub ExcluirGrupoProduto(lngGRUPO_PRODUTOID As Long)
  On Error GoTo trata
  Dim objGrupoProduto As datSisMetal.clsGrupoProduto
  Set objGrupoProduto = New datSisMetal.clsGrupoProduto
  
  'Excluir GrupoProduto
  objGrupoProduto.ExcluirGrupoProduto lngGRUPO_PRODUTOID

  '
  Set objGrupoProduto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub


