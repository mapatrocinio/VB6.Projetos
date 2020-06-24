VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVeiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirVeiculo(lngVEICULOID As Long, _
                         strPlaca As String, _
                         lngMODELOID As Long, _
                         strAno As String, _
                         strObservacao As String)

  On Error GoTo trata
  Dim strSql  As String
  Dim objRs   As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO VEICULO (PLACA, MODELOID, ANO, OBSERVACAO) VALUES (" & _
    Formata_Dados(strPlaca, tpDados_Texto, tpNulo_NaoAceita) & _
    ", " & Formata_Dados(IIf(lngMODELOID = 0, "", lngMODELOID), tpDados_Longo) & _
    ", " & Formata_Dados(strAno, tpDados_Texto, tpNulo_NaoAceita) & _
    ", " & Formata_Dados(strObservacao, tpDados_Texto, tpNulo_NaoAceita) & _
    ");"
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngVEICULOID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsVeiculo.InserirVeiculo]", _
            Err.Description
End Sub

Public Sub AlterarVeiculo(lngVEICULOID As Long, _
                         strPlaca As String, _
                         lngMODELOID As Long, _
                         strAno As String, _
                         strObservacao As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE VEICULO SET PLACA = " & Formata_Dados(strPlaca, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    " MODELOID = " & Formata_Dados(lngMODELOID, tpDados_Longo) & ", " & _
    " ANO = " & Formata_Dados(strAno, tpDados_Texto) & ", " & _
    " OBSERVACAO = " & Formata_Dados(strObservacao, tpDados_Texto) & " " & _
    " WHERE PKID = " & lngVEICULOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsVeiculo.AlterarVeiculo]", _
            Err.Description
End Sub

Public Function ListarVeiculo(lngVEICULOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT VEICULO.*, MARCA.NOME + '/' + MODELO.NOME AS MARCA_MODELO " & _
    "FROM VEICULO " & _
    " LEFT JOIN MODELO ON MODELO.PKID = VEICULO.MODELOID " & _
    " LEFT JOIN MARCA ON MARCA.PKID = MODELO.MARCAID " & _
    "WHERE VEICULO.PKID = " & _
    Formata_Dados(lngVEICULOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarVeiculo = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsVeiculo.ListarVeiculo]", _
            Err.Description
End Function

Public Sub ExcluirVeiculo(lngVEICULOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui VEICULO
  strSql = "DELETE FROM VEICULO WHERE VEICULO.PKID = " & _
    Formata_Dados(lngVEICULOID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsVeiculo.ExcluirVeiculo]", _
            Err.Description
End Sub


