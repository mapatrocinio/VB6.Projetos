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
  '
  Dim strSql As String
  
  strSql = "SELECT GRUPO_PRODUTO.*, REFERENCIAPRODUTO.REFERENCIA AS DESC_REFERENCIA " & _
    " FROM GRUPO_PRODUTO " & _
    " LEFT JOIN REFERENCIAPRODUTO ON REFERENCIAPRODUTO.PKID = GRUPO_PRODUTO.REFERENCIAPRODUTOID " & _
    " WHERE GRUPO_PRODUTO.PKID = " & _
    Formata_Dados(lngGRUPO_PRODUTOID, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarGrupoProduto = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsGrupoProduto.SelecionarGrupoProduto]"
End Function
Public Sub ExcluirGrupoProduto(lngGRUPO_PRODUTOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM GRUPO_PRODUTO " & _
    " WHERE PKID = " & lngGRUPO_PRODUTOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsGrupoProduto.ExcluirGrupoProduto]"
End Sub

Public Sub AlterarGrupoProduto(strNome As String, _
                        lngREFERENCIAPRODUTOID As Long, _
                        lngGRUPO_PRODUTOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE GRUPO_PRODUTO SET " & _
    " NOME = " & Formata_Dados(strNome, tpDados_Texto) & _
    ", REFERENCIAPRODUTOID = " & Formata_Dados(lngREFERENCIAPRODUTOID, tpDados_Longo) & _
    " WHERE PKID = " & Formata_Dados(lngGRUPO_PRODUTOID, tpDados_Longo)
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsGrupoProduto.AlterarGrupoProduto]"
End Sub
Public Sub IncluirGrupoProduto(strNome As String, _
                        lngREFERENCIAPRODUTOID As Long)
  On Error GoTo trata
  '
  Dim strSql          As String
  '
  strSql = "INSERT INTO GRUPO_PRODUTO (" & _
    " NOME, REFERENCIAPRODUTOID) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto) & _
    ", " & Formata_Dados(lngREFERENCIAPRODUTOID, tpDados_Longo) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsGrupoProduto.IncluirGrupoProduto]"
End Sub

