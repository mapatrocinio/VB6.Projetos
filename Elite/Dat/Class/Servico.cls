VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public Function ExcluirServico(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  'Delete SERVICO
  strSql = "DELETE FROM SERVICO " & _
    "WHERE SERVICO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsServico.ExcluirServico]", _
            Err.Description
End Function

Public Function SelecionarPacoteServicoPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PACOTESERVICO.* " & _
    "  " & _
    " FROM PACOTESERVICO " & _
    " LEFT JOIN AGENCIACNPJ ON AGENCIACNPJ.PKID = SERVICO.AGENCIACNPJID " & _
    " LEFT JOIN AGENCIA ON AGENCIA.PKID = AGENCIACNPJ.AGENCIAID " & _
    "WHERE PACOTESERVICO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarPacoteServicoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsServico.SelecionarPacoteServicoPeloPkid]", _
            Err.Description
End Function

Public Function SelecionarServicoPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT SERVICO.*, AGENCIA.NOME + ' (' + dbo.formataCNPJ(AGENCIACNPJ.CNPJ) + ')'  AS DESC_AGENCIA, " & _
    " ORIGEM.NOME AS DESC_ORIGEM, DESTINO.NOME AS DESC_DESTINO " & _
    " FROM SERVICO " & _
    " LEFT JOIN AGENCIACNPJ ON AGENCIACNPJ.PKID = SERVICO.AGENCIACNPJID " & _
    " LEFT JOIN AGENCIA ON AGENCIA.PKID = AGENCIACNPJ.AGENCIAID " & _
    " LEFT JOIN ORIGEM ON ORIGEM.PKID = SERVICO.ORIGEMID " & _
    " LEFT JOIN ORIGEM DESTINO ON DESTINO.PKID = SERVICO.DESTINOID " & _
    "WHERE SERVICO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarServicoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsServico.SelecionarServicoPeloPkid]", _
            Err.Description
End Function

Public Sub AlterarServico(lngPKID As Long, _
                          strDataHora As String, _
                          lngAgenciaCNPJId As Long, _
                          lngORIGEMID As Long, _
                          lngDESTINOID As Long, _
                          strSolicitante As String, _
                          strPassageiro As String, _
                          strReserva As String, _
                          strQtdPassageiro As String, _
                          strQtdCrianca As String, _
                          strVoo As String, _
                          strTerminal As String, _
                          strValor As String, _
                          strObservacao As String, _
                          strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE SERVICO SET " & _
    "DATAHORA = " & Formata_Dados(strDataHora, tpDados_DataHora) & ", " & _
    "AGENCIACNPJID = " & Formata_Dados(lngAgenciaCNPJId, tpDados_Longo) & ", " & _
    "ORIGEMID = " & Formata_Dados(lngORIGEMID, tpDados_Longo) & ", " & _
    "DESTINOID = " & Formata_Dados(lngDESTINOID, tpDados_Longo) & ", " & _
    "SOLICITANTE = " & Formata_Dados(strSolicitante, tpDados_Texto) & ", " & _
    "PASSAGEIRO = " & Formata_Dados(strPassageiro, tpDados_Texto) & ", "
  strSql = strSql & "RESERVA = " & Formata_Dados(strReserva, tpDados_Texto) & ", " & _
    "QTDPASSAGEIRO = " & Formata_Dados(strQtdPassageiro, tpDados_Longo) & ", " & _
    "QTDCRIANCAS = " & Formata_Dados(strQtdCrianca, tpDados_Longo) & ", " & _
    "VOO = " & Formata_Dados(strVoo, tpDados_Texto) & ", " & _
    "TERMINAL = " & Formata_Dados(strTerminal, tpDados_Texto) & ", " & _
    "VALOR = " & Formata_Dados(strValor, tpDados_Moeda) & ", " & _
    "OBSERVACAO = " & Formata_Dados(strObservacao, tpDados_Texto) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsServico.AlterarServico]", _
            Err.Description
End Sub

Public Sub AlterarVeiculo(lngPACOTESERVICOID As Long, _
                          lngVEICULOID As Long)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE PACOTESERVICO SET " & _
    "VEICULOID = " & Formata_Dados(lngVEICULOID, tpDados_Longo) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPACOTESERVICOID, tpDados_Longo)
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsServico.AlterarVeiculo]", _
            Err.Description
End Sub

Public Sub AlterarStatusServico(lngPKID As Long, _
                                strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE SERVICO SET " & _
   "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsServico.AlterarStatusServico]", _
            Err.Description
End Sub

Public Sub InserirServico(lngPKID As Long, _
                          strDataHora As String, _
                          lngAgenciaCNPJId As Long, _
                          lngORIGEMID As Long, _
                          lngDESTINOID As Long, _
                          strSolicitante As String, _
                          strPassageiro As String, _
                          strReserva As String, _
                          strQtdPassageiro As String, _
                          strQtdCrianca As String, _
                          strVoo As String, _
                          strTerminal As String, _
                          strValor As String, _
                          strObservacao As String, _
                          strStatus As String)

  On Error GoTo trata
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO SERVICO (DATAHORA, AGENCIACNPJID, ORIGEMID, DESTINOID, SOLICITANTE, PASSAGEIRO, " & _
    "RESERVA, QTDPASSAGEIRO, QTDCRIANCAS, VOO, TERMINAL, VALOR, OBSERVACAO, STATUS) VALUES (" & _
    Formata_Dados(strDataHora, tpDados_DataHora) & ", " & _
    Formata_Dados(lngAgenciaCNPJId, tpDados_Longo) & ", " & _
    Formata_Dados(lngORIGEMID, tpDados_Longo) & ", " & _
    Formata_Dados(lngDESTINOID, tpDados_Longo) & ", " & _
    Formata_Dados(strSolicitante, tpDados_Texto) & ", " & _
    Formata_Dados(strPassageiro, tpDados_Texto) & ", "
  strSql = strSql & Formata_Dados(strReserva, tpDados_Texto) & ", " & _
    Formata_Dados(strQtdPassageiro, tpDados_Longo) & ", " & _
    Formata_Dados(strQtdCrianca, tpDados_Longo) & ", " & _
    Formata_Dados(strVoo, tpDados_Texto) & ", " & _
    Formata_Dados(strTerminal, tpDados_Texto) & ", " & _
    Formata_Dados(strValor, tpDados_Moeda) & ", " & _
    Formata_Dados(strObservacao, tpDados_Texto) & ", " & _
    Formata_Dados(strStatus, tpDados_Texto) & ");"
  '
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngPKID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsServico.InserirServico]", _
            Err.Description
End Sub

Public Sub DesativarServicoDoPacote(lngPACOTESERVICOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  '
  strSql = "UPDATE PACOTESERVICO SET " & _
    "STATUS = " & Formata_Dados("C", tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPACOTESERVICOID, tpDados_Longo)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoCortesia.DesativarServicoDoPacote]", _
            Err.Description
End Sub

Public Sub DesassociarServicoAoPacote(lngPACOTESERVICOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  '
  strSql = "DELETE FROM PACOTESERVICO WHERE " & _
    " PACOTESERVICO.PKID = " & Formata_Dados(lngPACOTESERVICOID, tpDados_Longo)

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoCortesia.DesassociarServicoAoPacote]", _
            Err.Description
End Sub

Public Sub AssociarServicoAoPacote(lngPACOTEID As Long, _
                                   lngSERVICOID As Long, _
                                   lngVEICULOID As Long, _
                                   strStatus As String)
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT * FROM PACOTESERVICO " & _
    " WHERE PACOTEID = " & Formata_Dados(lngPACOTEID, tpDados_Longo) & _
    " AND SERVICOID = " & Formata_Dados(lngSERVICOID, tpDados_Longo) & _
    ";"
  strSql = strSql & vbCrLf & " IF @@ROWCOUNT > 0 "
  strSql = strSql & vbCrLf & " UPDATE PACOTESERVICO SET STATUS = " & Formata_Dados("A", tpDados_Texto) & ", " & _
    " VEICULOID = " & Formata_Dados(lngVEICULOID, tpDados_Longo) & " " & _
    " WHERE PACOTEID = " & Formata_Dados(lngPACOTEID, tpDados_Longo) & _
    " AND SERVICOID = " & Formata_Dados(lngSERVICOID, tpDados_Longo)
  strSql = strSql & vbCrLf & " ELSE "
  strSql = strSql & vbCrLf & " INSERT INTO PACOTESERVICO (PACOTEID, SERVICOID, VEICULOID, STATUS) VALUES (" & _
    Formata_Dados(lngPACOTEID, tpDados_Longo) & ", " & _
    Formata_Dados(lngSERVICOID, tpDados_Longo) & ", " & _
    Formata_Dados(lngVEICULOID, tpDados_Longo) & ", " & _
    Formata_Dados(strStatus, tpDados_Texto) & _
    ")"
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTipoCortesia.AssociarServicoAoPacote]", _
            Err.Description
End Sub



