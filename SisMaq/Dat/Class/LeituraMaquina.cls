VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLeituraMaquina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirLeituraMaquina(lngLEITURAMAQUINAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM LEITURAMAQUINA " & _
    "WHERE LEITURAMAQUINA.PKID = " & _
    Formata_Dados(lngLEITURAMAQUINAID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLeituraMaquina.ExcluirLeituraMaquina]", _
            Err.Description
End Function


Public Function SelecionarLeituraMaquinaPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT LEITURAMAQUINA.* " & _
    "FROM LEITURAMAQUINA " & _
    "WHERE LEITURAMAQUINA.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  Set SelecionarLeituraMaquinaPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLeituraMaquina.SelecionarLeituraMaquinaPeloPkid]", _
            Err.Description
End Function

Public Sub TratarLeituraMaquinaStatus(lngLEITURAMAQUINAID As Long, _
                                      lngMAQUINAID As Long, _
                                      strData As String, _
                                      strAtualMedicaoEntrada As String, _
                                      strAtualMedicaoSaida As String)

  On Error GoTo trata
  Dim strStatus As String
  Dim lngLEITURAMAQUINAANTID As Long
  Dim curMedicaoEntrada       As Currency
  Dim curMedicaoSaida         As Currency
  '
  Dim strSql  As String
  Dim objRs   As ADODB.Recordset
  '
  strSql = "SELECT " & _
        " LEITURAMAQUINA.PKID, " & _
        " LEITURAMAQUINA.MEDICAOENTRADA, " & _
        " LEITURAMAQUINA.MEDICAOSAIDA " & _
        " FROM LEITURAMAQUINA " & _
        " WHERE LEITURAMAQUINA.LEITURAID IN " & _
        "       (SELECT PKID " & _
        "       FROM LEITURA " & _
        "       WHERE DATA = " & _
        "             (SELECT MAX(DATA) FROM LEITURA " & _
        "             WHERE LEITURA.DATA < " & Formata_Dados(strData, tpDados_DataHora) & ")) " & _
        " AND LEITURAMAQUINA.MAQUINAID = " & Formata_Dados(lngMAQUINAID, tpDados_Longo)
  '
  Set objRs = RunSPReturnRS(strSql)
  lngLEITURAMAQUINAANTID = 0
  If Not objRs.EOF Then
    lngLEITURAMAQUINAANTID = objRs.Fields("PKID").Value
    curMedicaoEntrada = objRs.Fields("MEDICAOENTRADA").Value
    curMedicaoSaida = objRs.Fields("MEDICAOSAIDA").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  If lngLEITURAMAQUINAANTID > 0 Then
    'Encontrou medicao anterior, compara com lançada
    If CCur(strAtualMedicaoEntrada) < curMedicaoEntrada Or _
      CCur(strAtualMedicaoSaida) < curMedicaoSaida Then
      'Uma das medições é maior que a anterior, loga como erro
      strSql = "UPDATE LEITURAMAQUINA SET " & _
              "STATUS = " & Formata_Dados("E", tpDados_Texto) & _
              ", LEITURAMAQUINAANTID = " & Formata_Dados(lngLEITURAMAQUINAANTID, tpDados_Longo) & _
              " WHERE PKID = " & Formata_Dados(lngLEITURAMAQUINAID, tpDados_Longo)
      RunSP strSql
    End If
  End If
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLeituraMaquina.TratarLeituraMaquinaStatus]", _
            Err.Description
End Sub

Public Sub InserirLeituraMaquina(lngPKID As Long, _
                                 lngLEITURAID As Long, _
                                 lngMAQUINAID As Long, _
                                 strMedicaoEntrada As String, _
                                 strMedicaoSaida As String)

  On Error GoTo trata
  Dim strSql  As String
  Dim objRs   As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO LEITURAMAQUINA (LEITURAID, MAQUINAID, MEDICAOENTRADA, MEDICAOSAIDA, STATUS) VALUES (" & _
    Formata_Dados(lngLEITURAID, tpDados_Longo) & _
    ", " & Formata_Dados(lngMAQUINAID, tpDados_Longo) & _
    ", " & Formata_Dados(strMedicaoEntrada, tpDados_Moeda) & _
    ", " & Formata_Dados(strMedicaoSaida, tpDados_Moeda) & _
    ", " & Formata_Dados("A", tpDados_Texto) & _
    ");"
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngPKID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLeituraMaquina.InserirLeituraMaquina]", _
            Err.Description
End Sub





