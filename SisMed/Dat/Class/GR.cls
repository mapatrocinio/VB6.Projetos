VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarGRPeloPkid(lngGRID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  strSql = "SELECT GR.*, PRONTUARIO.NOME, PRONTUARIO.DTNASCIMENTO, ESPECIALIDADE.ESPECIALIDADE AS ESPEC_ESPECIALIDADE, SALA.NUMERO AS SALA_NUMERO,  ATENDE.HORAINICIO + ' - ' + ATENDE.HORATERMINO AS PERIODO_PERIODO, PRESTADOR.NOME AS PREST_PRESTADOR, FUNCIONARIO.NIVEL " & _
    "FROM GR INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = GR.PRONTUARIOID " & _
    " INNER JOIN TURNO ON TURNO.PKID = GR.TURNOID " & _
    " INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
      " INNER JOIN SALA ON SALA.PKID = ATENDE.SALAID " & _
      " INNER JOIN PRONTUARIO AS PRESTADOR ON PRESTADOR.PKID = ATENDE.PRONTUARIOID " & _
    " INNER JOIN ESPECIALIDADE ON ESPECIALIDADE.PKID = GR.ESPECIALIDADEID " & _
    " INNER JOIN FUNCIONARIO ON GR.FUNCIONARIOID = FUNCIONARIO.PRONTUARIOID " & _
    "WHERE GR.PKID = " & _
    Formata_Dados(lngGRID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarGRPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGR.SelecionarGRPeloPkid]", _
            Err.Description
End Function

Public Function CapturaPrestEspec(strCodigo As String, _
                                  strDiaDaSemana As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  Dim strData As String
  strData = Formata_Dados("01/01/1900 " & Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00"), tpDados_DataHora)
      
  strSql = "SELECT PRONTUARIO.NOME AS PREST_PRESTADOR, ESPECIALIDADE.ESPECIALIDADE AS ESPEC_ESPECIALIDADE, SALA.NUMERO AS SALA_NUMERO,  ATENDE.HORAINICIO + ' - ' + ATENDE.HORATERMINO as PERIODO_PERIODO " & _
    "FROM PRONTUARIO INNER JOIN ATENDE ON PRONTUARIO.PKID = ATENDE.PRONTUARIOID " & _
      " INNER JOIN SALA ON SALA.PKID = ATENDE.SALAID " & _
    " INNER JOIN PRESTADORESPECIALIDADE ON PRONTUARIO.PKID = PRESTADORESPECIALIDADE.PRONTUARIOID " & _
    " INNER JOIN ESPECIALIDADE ON ESPECIALIDADE.PKID = PRESTADORESPECIALIDADE.ESPECIALIDADEID " & _
    " INNER JOIN DIASDASEMANA ON DIASDASEMANA.PKID = ATENDE.DIASDASEMANAID " & _
    "WHERE (ESPECIALIDADE.ESPECIALIDADE LIKE " & Formata_Dados(strCodigo & "%", tpDados_Texto, tpNulo_NaoAceita) & _
    " OR PRONTUARIO.NOME LIKE " & Formata_Dados(strCodigo & "%", tpDados_Texto, tpNulo_NaoAceita) & ")"
    '" AND convert(datetime, '01/01/1900 ' + HORAINICIO ,103) <= " & strData & _
    '" AND convert(datetime, '01/01/1900 ' + HORATERMINO ,103) >= " & strData
  strSql = strSql & " AND DIASDASEMANA.DIADASEMANA = " & Formata_Dados(strDiaDaSemana, tpDados_Texto) & _
    " AND ATENDE.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
    " AND SALA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
    " ORDER BY PRONTUARIO.NOME, ESPECIALIDADE.ESPECIALIDADE, SALA.NUMERO, ATENDE.HORAINICIO, ATENDE.HORATERMINO"
            
  Set CapturaPrestEspec = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGR.CapturaPrestEspec]", _
            Err.Description
End Function


Public Function CapturaProcedimento(strProcedimento As String, _
                                    lngPRESTADORID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  Dim strData As String
  strSql = "SELECT PROCEDIMENTO.INDACEITAVALOR, PROCEDIMENTO.PROCEDIMENTO, PROCEDIMENTO.VALOR AS VALOR_PROC, PROCEDIMENTO.* " & _
    "FROM PROCEDIMENTO INNER JOIN PRESTADORPROCEDIMENTO ON PROCEDIMENTO.PKID = PRESTADORPROCEDIMENTO.PROCEDIMENTOID " & _
    "WHERE PROCEDIMENTO.PROCEDIMENTO LIKE " & Formata_Dados(strProcedimento & "%", tpDados_Texto, tpNulo_NaoAceita) & _
    " AND PRESTADORPROCEDIMENTO.PRONTUARIOID = " & Formata_Dados(lngPRESTADORID, tpDados_Longo) & _
    " ORDER BY PROCEDIMENTO.PROCEDIMENTO"
            
  Set CapturaProcedimento = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGR.CapturaProcedimento]", _
            Err.Description
End Function

Public Function CapturaProcedimentoGeral(strProcedimento As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  Dim strData As String
  strSql = "SELECT PROCEDIMENTO.INDACEITAVALOR, PROCEDIMENTO.PROCEDIMENTO, PROCEDIMENTO.VALOR AS VALOR_PROC, PROCEDIMENTO.* " & _
    "FROM PROCEDIMENTO "
    
  If strProcedimento & "" <> "" Then
    strSql = strSql & "WHERE PROCEDIMENTO.PROCEDIMENTO LIKE " & Formata_Dados("%" & strProcedimento & "%", tpDados_Texto, tpNulo_NaoAceita)
  End If
    strSql = strSql & " ORDER BY PROCEDIMENTO.PROCEDIMENTO"
  Set CapturaProcedimentoGeral = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGR.CapturaProcedimentoGeral]", _
            Err.Description
End Function


Public Function CapturaProntuario(strNome As String, _
                                  strCPF As String, _
                                  strDtNascimento) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql    As String
  Dim strWhere  As String
  
  strWhere = "  WHERE "
  If Len(strNome & "") <> 0 Then
    strWhere = strWhere & " PRONTUARIO.NOME LIKE " & Formata_Dados(strNome & "%", tpDados_Texto)
  End If
  If Len(strCPF & "") <> 0 Then
    If strWhere <> "  WHERE " Then strWhere = strWhere & " AND "
    strWhere = strWhere & " PRONTUARIO.CPF = " & Formata_Dados(strCPF, tpDados_Texto)
  End If
  If Len(strDtNascimento & "") <> 0 Then
    If strWhere <> "  WHERE " Then strWhere = strWhere & " AND "
    strWhere = strWhere & " PRONTUARIO.DTNASCIMENTO = " & Formata_Dados(strDtNascimento, tpDados_DataHora)
  End If
  
  
  strSql = "SELECT PRONTUARIO.NOME, PRONTUARIO.CPF, PRONTUARIO.DTNASCIMENTO " & _
    "FROM PRONTUARIO " & _
    strWhere & _
    " ORDER BY PRONTUARIO.NOME, PRONTUARIO.CPF, PRONTUARIO.DTNASCIMENTO"
            
  Set CapturaProntuario = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGR.CapturaProntuario]", _
            Err.Description
End Function


Public Sub AlterarGR(lngPKID As Long, _
                     lngPRONTUARIOID As Long, _
                     lngAtendeId As Long, _
                     lngESPECIALIDADEID As Long, _
                     strDescricao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE GR SET " & _
    "PRONTUARIOID = " & Formata_Dados(lngPRONTUARIOID, tpDados_Longo) & ", " & _
    "ATENDEID = " & Formata_Dados(lngAtendeId, tpDados_Longo) & ", " & _
    "ESPECIALIDADEID = " & Formata_Dados(lngESPECIALIDADEID, tpDados_Longo) & ", " & _
    "DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGR.AlterarGR]", _
            Err.Description
End Sub

Public Sub AlterarGRNaoAtendida(lngPKID As Long, _
                                strDataNaoAtendimento As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE GR SET " & _
    "DATANAOATENDIDA = " & Formata_Dados(strDataNaoAtendimento, tpDados_DataHora) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGR.AlterarGRNaoAtendida]", _
            Err.Description
End Sub

Public Sub AlterarGRLiberarAtend(lngPKID As Long, _
                                 strDataLiberaAtend As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE GR SET " & _
    "DATALIBERAATEND = " & Formata_Dados(strDataLiberaAtend, tpDados_DataHora) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGR.AlterarGRLiberarAtend]", _
            Err.Description
End Sub

Public Sub AlterarGRExpirarAtend(lngPKID As Long, _
                                 strDataExpiraAtend As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE GR SET " & _
    "DATACONFIRMAEXPIRACAO = " & Formata_Dados(strDataExpiraAtend, tpDados_DataHora) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGR.AlterarGRExpirarAtend]", _
            Err.Description
End Sub

Public Sub InserirGR(ByRef lngPKID As Long, _
                     lngPRONTUARIOID As Long, _
                     lngTurnoId As Long, _
                     lngTurnoLabId As String, _
                     lngAtendeId As Long, _
                     lngESPECIALIDADEID As Long, _
                     strSequencial As String, _
                     strSequencialSenha As String, _
                     strData As String, _
                     strStatus As String, _
                     strStatusImp As String, _
                     lngFuncionarioId As Long, _
                     strCortesia As String, _
                     strUsuLib As String, _
                     strDescricao As String)
  On Error GoTo trata
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO GR (PRONTUARIOID, TURNOID, TURNOLABID, ATENDEID, ESPECIALIDADEID, SEQUENCIAL, SENHA, DATA, STATUS, STATUSIMPRESSAO, FUNCIONARIOID, INDCORTESIA, USUARIOCORT, DESCRICAO) VALUES (" & _
    " " & Formata_Dados(lngPRONTUARIOID, tpDados_Longo) & _
    ", " & Formata_Dados(lngTurnoId, tpDados_Longo) & _
    ", " & Formata_Dados(lngTurnoLabId, tpDados_Longo) & _
    ", " & Formata_Dados(lngAtendeId, tpDados_Longo) & _
    ", " & Formata_Dados(lngESPECIALIDADEID, tpDados_Longo) & _
    ", " & Formata_Dados(strSequencial, tpDados_Longo) & _
    ", " & Formata_Dados(strSequencialSenha, tpDados_Longo) & _
    ", " & Formata_Dados(strData, tpDados_DataHora) & _
    ", " & Formata_Dados(strStatus, tpDados_Texto) & _
    ", " & Formata_Dados(strStatusImp, tpDados_Texto) & _
    ", " & Formata_Dados(lngFuncionarioId, tpDados_Longo) & _
    ", " & Formata_Dados(strCortesia, tpDados_Texto) & _
    ", " & Formata_Dados(strUsuLib, tpDados_Texto) & _
    ", " & Formata_Dados(strDescricao, tpDados_Texto) & ");"
    
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
            Err.Source & ".[clsGR.InserirGR]", _
            Err.Description
End Sub


Public Sub InserirGRPROCEDIMENTO(lngGRID As Long, _
                                 lngPROCEDIMENTOID As Long, _
                                 strQuantidade As String, _
                                 strValor As String, _
                                 strValorCort As String)
  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO GRPROCEDIMENTO (GRID, PROCEDIMENTOID, QTD, VALOR, VALORCORTESIA) VALUES (" & _
    " " & Formata_Dados(lngGRID, tpDados_Longo) & _
    ", " & Formata_Dados(lngPROCEDIMENTOID, tpDados_Longo) & _
    ", " & Formata_Dados(strQuantidade, tpDados_Longo) & _
    ", " & Formata_Dados(strValor, tpDados_Moeda) & _
    ", " & Formata_Dados(strValorCort, tpDados_Moeda) & ");"
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGRPROCEDIMENTO.InserirGRPROCEDIMENTO]", _
            Err.Description
End Sub


Public Function ExcluirGRPROCEDIMENTO(lngGRPROCEDIMENTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "DELETE FROM GRPROCEDIMENTO " & _
    "WHERE PKID = " & _
    Formata_Dados(lngGRPROCEDIMENTOID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGRPROCEDIMENTO.ExcluirGRPROCEDIMENTO]", _
            Err.Description
End Function


Public Function SelecionarGRTotalProc(lngGRID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  strSql = "SELECT isnull(SUM(GRPROCEDIMENTO.VALOR), 0) AS TOTALAPAGAR " & _
    "FROM GRPROCEDIMENTO " & _
    "WHERE GRPROCEDIMENTO.GRID = " & _
    Formata_Dados(lngGRID, tpDados_Longo)
  Set SelecionarGRTotalProc = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGR.SelecionarGRTotalProc]", _
            Err.Description
End Function

Public Sub AlterarStatusGR(lngGRID As Long, _
                           strStatus As String, _
                           strStatusImpressao As String, _
                           Optional lngTurnoId As Long, _
                           Optional gsNomeUsuLib As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE GR SET "
  
  If Len(strStatus & "") > 0 Then
    strSql = strSql & " STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " "
    If lngTurnoId <> 0 Then
      strSql = strSql & ", TURNOCANCID = " & Formata_Dados(lngTurnoId, tpDados_Longo) & " "
    End If
    If Len(gsNomeUsuLib & "") > 0 Then
      strSql = strSql & ", USUARIOLIB = " & Formata_Dados(gsNomeUsuLib, tpDados_Texto) & " "
    End If
  End If
  
  If Len(strStatusImpressao & "") > 0 Then
    If strSql <> "UPDATE GR SET " Then
      strSql = strSql & " , "
    End If
    strSql = strSql & " STATUSIMPRESSAO = " & Formata_Dados(strStatusImpressao, tpDados_Texto) & " "
  End If
  strSql = strSql & " WHERE PKID = " & Formata_Dados(lngGRID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGR.AlterarStatusGR]", _
            Err.Description
End Sub

Public Sub AlterarArquivoGR(lngGRID As Long, _
                            strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE GR SET ARQUIVO = " & Formata_Dados(strStatus, tpDados_Texto)
  strSql = strSql & " WHERE PKID = " & Formata_Dados(lngGRID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGR.AlterarArquivoGR]", _
            Err.Description
End Sub

Public Sub AlterarTurnoGR(lngGRID As Long, _
                          lngTurnoId As Long, _
                          strFuncionarioId As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE GR SET " & _
      " TURNOID = " & Formata_Dados(lngTurnoId, tpDados_Longo) & _
      ", FUNCIONARIOFECHAID = " & Formata_Dados(strFuncionarioId, tpDados_Longo)
  strSql = strSql & " WHERE PKID = " & Formata_Dados(lngGRID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGR.AlterarTurnoGR]", _
            Err.Description
End Sub


Public Function CapturaGRTurnoCorrente(strTURNOID As String, _
                                       lngPRONTUARIOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  Dim strData As String
  strSql = "SELECT DISTINCT GR.PKID, GR.DATA, GR.SEQUENCIAL, PRONTUARIO.NOME, ESPECIALIDADE.ESPECIALIDADE, PRESTADOR.NOME, SALA.NUMERO AS SALA_NUMERO,  '(' + FUNC.NIVEL + ') ' + FUNCIONARIO.NOME, GR.STATUS, GR.STATUSIMPRESSAO " & _
    "FROM GR INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = GR.PRONTUARIOID " & _
    " INNER JOIN TURNO ON TURNO.PKID = GR.TURNOID " & _
    "   OR TURNO.PKID = GR.TURNOLABID " & _
    " INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
      " INNER JOIN SALA ON SALA.PKID = ATENDE.SALAID " & _
      " INNER JOIN PRONTUARIO AS PRESTADOR ON PRESTADOR.PKID = ATENDE.PRONTUARIOID " & _
    " INNER JOIN PRONTUARIO AS FUNCIONARIO ON FUNCIONARIO.PKID = GR.FUNCIONARIOID " & _
    " INNER JOIN FUNCIONARIO AS FUNC ON FUNCIONARIO.PKID = FUNC.PRONTUARIOID " & _
    " INNER JOIN ESPECIALIDADE ON ESPECIALIDADE.PKID = GR.ESPECIALIDADEID " & _
    "WHERE GR.TURNOID IN " & _
      strTURNOID & _
      " OR (TURNO.STATUS = " & Formata_Dados(1, tpDados_Longo) & _
      " AND TURNO.PRONTUARIOID = " & Formata_Dados(lngPRONTUARIOID, tpDados_Longo) & ")" & _
      " ORDER BY GR.DATA DESC, GR.SEQUENCIAL DESC "
  Set CapturaGRTurnoCorrente = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGR.CapturaGRTurnoCorrente]", _
            Err.Description
End Function

Public Function CapturaGRTurnoCorrenteTODOS(strTURNOID As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  Dim strData As String
  strSql = "SELECT DISTINCT GR.PKID, GR.DATA, GR.SEQUENCIAL, PRONTUARIO.NOME, ESPECIALIDADE.ESPECIALIDADE, PRESTADOR.NOME, SALA.NUMERO AS SALA_NUMERO,  '(' + FUNC.NIVEL + ') ' + FUNCIONARIO.NOME, GR.STATUS, GR.STATUSIMPRESSAO " & _
    "FROM GR INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = GR.PRONTUARIOID " & _
    " INNER JOIN TURNO ON TURNO.PKID = GR.TURNOID " & _
    "   OR TURNO.PKID = GR.TURNOLABID " & _
    " INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
      " INNER JOIN SALA ON SALA.PKID = ATENDE.SALAID " & _
      " INNER JOIN PRONTUARIO AS PRESTADOR ON PRESTADOR.PKID = ATENDE.PRONTUARIOID " & _
    " INNER JOIN PRONTUARIO AS FUNCIONARIO ON FUNCIONARIO.PKID = GR.FUNCIONARIOID " & _
    " INNER JOIN FUNCIONARIO AS FUNC ON FUNCIONARIO.PKID = FUNC.PRONTUARIOID " & _
    " INNER JOIN ESPECIALIDADE ON ESPECIALIDADE.PKID = GR.ESPECIALIDADEID " & _
    "WHERE GR.TURNOID IN " & _
      strTURNOID & _
      " ORDER BY GR.DATA DESC, GR.SEQUENCIAL DESC "
  Set CapturaGRTurnoCorrenteTODOS = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGR.CapturaGRTurnoCorrenteTODOS]", _
            Err.Description
End Function

Public Function CapturaGRTurnoCorrenteARQ(strTURNOID As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  Dim strData As String
  strSql = "SELECT DISTINCT GR.PKID, GR.DATA, GR.SEQUENCIAL, PRONTUARIO.NOME, PRONTUARIO.DTNASCIMENTO, ESPECIALIDADE.ESPECIALIDADE, PRESTADOR.NOME, SALA.NUMERO AS SALA_NUMERO,  '(' + FUNC.NIVEL + ') ' + FUNCIONARIO.NOME, PRONTUARIO.PKID, GR.ARQUIVO, GR.STATUSIMPRESSAO, GR.STATUS " & _
    "FROM GR INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = GR.PRONTUARIOID " & _
    " INNER JOIN TURNO ON TURNO.PKID = GR.TURNOID " & _
    "   OR TURNO.PKID = GR.TURNOLABID " & _
    " INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
      " INNER JOIN SALA ON SALA.PKID = ATENDE.SALAID " & _
      " INNER JOIN PRONTUARIO AS PRESTADOR ON PRESTADOR.PKID = ATENDE.PRONTUARIOID " & _
    " INNER JOIN PRONTUARIO AS FUNCIONARIO ON FUNCIONARIO.PKID = GR.FUNCIONARIOID " & _
    " INNER JOIN FUNCIONARIO AS FUNC ON FUNCIONARIO.PKID = FUNC.PRONTUARIOID " & _
    " INNER JOIN ESPECIALIDADE ON ESPECIALIDADE.PKID = GR.ESPECIALIDADEID " & _
    "WHERE GR.TURNOID IN " & _
      strTURNOID & _
      " AND TURNOLABID IS NULL " & _
      " ORDER BY GR.DATA DESC, GR.SEQUENCIAL DESC "
  Set CapturaGRTurnoCorrenteARQ = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGR.CapturaGRTurnoCorrenteARQ]", _
            Err.Description
End Function



Public Function CapturaGRTurnoCorrenteFIN(strDataIniAtual As String, _
                                          strDataFimAtual As String, _
                                          intMaxDiasAtend As Integer) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  Dim strData As String
  'REGRA: GR´s fechadas na data atual
  'ou GR´s COM STATUS IN (NÃO ATENDIDA LIBERADA PARA ATENDIMENTO)
  'ou GR´s COM STATUS = EXPIRADA E EXPIRADA NA DATA ATUAL
  strSql = "SELECT DISTINCT GR.PKID, GR.DATA, GR.SEQUENCIAL, PRONTUARIO.NOME, PRONTUARIO.DTNASCIMENTO, ESPECIALIDADE.ESPECIALIDADE, PRESTADOR.NOME, SALA.NUMERO AS SALA_NUMERO,  '(' + FUNC.NIVEL + ') ' + FUNCIONARIO.NOME, PRONTUARIO.PKID, " & _
    " CASE " & _
    "   WHEN GR.STATUS = 'N' AND DATEDIFF ( d, GR.DATA , " & Formata_Dados(Format(Now, "DD/MM/YYYY hh:mm"), tpDados_DataHora) & " ) > " & Formata_Dados(intMaxDiasAtend, tpDados_Longo) & " THEN 'Z' " & _
    "   ELSE GR.STATUS END " & _
    "FROM GR INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = GR.PRONTUARIOID " & _
    " INNER JOIN TURNO ON TURNO.PKID = GR.TURNOID " & _
    "   OR TURNO.PKID = GR.TURNOLABID " & _
    " INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
      " INNER JOIN SALA ON SALA.PKID = ATENDE.SALAID " & _
      " INNER JOIN PRONTUARIO AS PRESTADOR ON PRESTADOR.PKID = ATENDE.PRONTUARIOID " & _
    " INNER JOIN PRONTUARIO AS FUNCIONARIO ON FUNCIONARIO.PKID = GR.FUNCIONARIOID " & _
    " INNER JOIN FUNCIONARIO AS FUNC ON FUNCIONARIO.PKID = FUNC.PRONTUARIOID " & _
    " INNER JOIN ESPECIALIDADE ON ESPECIALIDADE.PKID = GR.ESPECIALIDADEID " & _
    "WHERE (GR.DATA >= " & Formata_Dados(strDataIniAtual, tpDados_DataHora) & _
      " AND GR.DATA < " & Formata_Dados(strDataFimAtual, tpDados_DataHora) & _
      " AND GR.STATUS = " & Formata_Dados("F", tpDados_Texto) & ")" & _
      " OR (GR.STATUS IN ('N','L')) " & _
      " OR (GR.DATACONFIRMAEXPIRACAO >= " & Formata_Dados(strDataIniAtual, tpDados_DataHora) & _
      " AND GR.DATACONFIRMAEXPIRACAO < " & Formata_Dados(strDataFimAtual, tpDados_DataHora) & _
      " AND GR.STATUS = " & Formata_Dados("E", tpDados_Texto) & ")" & _
      " ORDER BY GR.DATA DESC, GR.SEQUENCIAL DESC "
  Set CapturaGRTurnoCorrenteFIN = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGR.CapturaGRTurnoCorrenteFIN]", _
            Err.Description
End Function







Public Function CapturaGRTurnoCorrentePRE(lngPRESTADORID As Long, _
                                          strDataIniAtual As String, _
                                          strDataFimAtual As String, _
                                          strNivel As String, _
                                          strPrestador As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  Dim strData As String
  If strNivel = "ARQ" Or strNivel = "ADM" Then
    strSql = "SELECT DISTINCT GR.PKID, ATENDIMENTO.INDSCANER, GR.DATA AS GR_DATA, GR.SEQUENCIAL AS GR_SEQUENCIAL, PRONTUARIO.NOME , PRONTUARIO.DTNASCIMENTO, ESPECIALIDADE.ESPECIALIDADE, PRESTADOR.NOME AS PRESTADOR_NOME, SALA.NUMERO AS SALA_NUMERO,  '(' + FUNC.NIVEL + ') ' + FUNCIONARIO.NOME, PRONTUARIO.PKID, GR.STATUS " & _
      "FROM GR INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = GR.PRONTUARIOID " & _
      " INNER JOIN TURNO ON TURNO.PKID = GR.TURNOID " & _
      "   OR TURNO.PKID = GR.TURNOLABID " & _
      " INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
        " INNER JOIN SALA ON SALA.PKID = ATENDE.SALAID " & _
        " INNER JOIN PRONTUARIO AS PRESTADOR ON PRESTADOR.PKID = ATENDE.PRONTUARIOID " & _
      " INNER JOIN PRONTUARIO AS FUNCIONARIO ON FUNCIONARIO.PKID = GR.FUNCIONARIOID " & _
      " INNER JOIN FUNCIONARIO AS FUNC ON FUNCIONARIO.PKID = FUNC.PRONTUARIOID " & _
      " INNER JOIN ESPECIALIDADE ON ESPECIALIDADE.PKID = GR.ESPECIALIDADEID " & _
      " LEFT JOIN ATENDIMENTO ON ATENDIMENTO.GRID = GR.PKID " & _
      "WHERE (((GR.STATUS = " & Formata_Dados("F", tpDados_Texto) & _
        " OR GR.STATUS = " & Formata_Dados("A", tpDados_Texto) & ")" & _
        " AND GR.DATA >= " & Formata_Dados(strDataIniAtual, tpDados_DataHora) & _
        " AND GR.DATA < " & Formata_Dados(strDataFimAtual, tpDados_DataHora) & ")" & _
        " OR (GR.STATUS = " & Formata_Dados("L", tpDados_Texto) & ")" & _
        " OR (GR.STATUS = " & Formata_Dados("P", tpDados_Texto) & _
        " AND ATENDIMENTO.DATA >= " & Formata_Dados(strDataIniAtual, tpDados_DataHora) & _
        " AND ATENDIMENTO.DATA < " & Formata_Dados(strDataFimAtual, tpDados_DataHora) & ")) " & _
        " AND ATENDIMENTO.INDSCANER = " & Formata_Dados("S", tpDados_Texto)
        
    If strPrestador & "" <> "" Then
      strSql = strSql & " AND PRESTADOR.NOME = " & Formata_Dados(strPrestador, tpDados_Texto)
    End If
    'Importadas
    strSql = strSql & " UNION ALL "
    strSql = strSql & "SELECT DISTINCT 0 AS GR_ID, ATENDIMENTO.INDSCANER, ATENDIMENTO.DATA AS GR_DATA, '' AS GR_SEQUENCIAL, PRONTUARIO.NOME, PRONTUARIO.DTNASCIMENTO, '' AS ESPECIALIDADE, '' AS PRESTADOR_NOME, '' AS SALA_NUMERO,  '' AS FUNCIONARIO_NOME, PRONTUARIO.PKID, '' AS GR_STATUS " & _
      "FROM ATENDIMENTO " & _
      " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = ATENDIMENTO.PRONTUARIOID " & _
      "WHERE ((ATENDIMENTO.DATA >= " & Formata_Dados(strDataIniAtual, tpDados_DataHora) & _
        " AND ATENDIMENTO.DATA < " & Formata_Dados(strDataFimAtual, tpDados_DataHora) & ")" & _
        " OR (ATENDIMENTO.DATAIMPORTACAO >= " & Formata_Dados(strDataIniAtual, tpDados_DataHora) & _
        " AND ATENDIMENTO.DATAIMPORTACAO < " & Formata_Dados(strDataFimAtual, tpDados_DataHora) & ")) " & _
        " AND ATENDIMENTO.INDSCANER = " & Formata_Dados("A", tpDados_Texto)
    '
    strSql = strSql & " ORDER BY GR_DATA DESC, GR_SEQUENCIAL DESC "

  ElseIf strNivel = "PRE" Then
    strSql = "SELECT DISTINCT GR.PKID, ATENDIMENTO.INDSCANER, GR.DATA, GR.SEQUENCIAL, PRONTUARIO.NOME , PRONTUARIO.DTNASCIMENTO, ESPECIALIDADE.ESPECIALIDADE, PRESTADOR.NOME AS PRESTADOR_NOME, SALA.NUMERO AS SALA_NUMERO,  '(' + FUNC.NIVEL + ') ' + FUNCIONARIO.NOME, PRONTUARIO.PKID, GR.STATUS " & _
      "FROM GR INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = GR.PRONTUARIOID " & _
      " INNER JOIN TURNO ON TURNO.PKID = GR.TURNOID " & _
      "   OR TURNO.PKID = GR.TURNOLABID " & _
      " INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
        " INNER JOIN SALA ON SALA.PKID = ATENDE.SALAID " & _
        " INNER JOIN PRONTUARIO AS PRESTADOR ON PRESTADOR.PKID = ATENDE.PRONTUARIOID " & _
      " INNER JOIN PRONTUARIO AS FUNCIONARIO ON FUNCIONARIO.PKID = GR.FUNCIONARIOID " & _
      " INNER JOIN FUNCIONARIO AS FUNC ON FUNCIONARIO.PKID = FUNC.PRONTUARIOID " & _
      " INNER JOIN ESPECIALIDADE ON ESPECIALIDADE.PKID = GR.ESPECIALIDADEID " & _
      " LEFT JOIN ATENDIMENTO ON ATENDIMENTO.GRID = GR.PKID " & _
      "WHERE (((GR.STATUS = " & Formata_Dados("F", tpDados_Texto) & _
        " OR GR.STATUS = " & Formata_Dados("A", tpDados_Texto) & ")" & _
        " AND GR.DATA >= " & Formata_Dados(strDataIniAtual, tpDados_DataHora) & _
        " AND GR.DATA < " & Formata_Dados(strDataFimAtual, tpDados_DataHora) & ")" & _
        " OR (GR.STATUS = " & Formata_Dados("L", tpDados_Texto) & ")" & _
        " OR (GR.STATUS = " & Formata_Dados("P", tpDados_Texto) & _
        " AND ATENDIMENTO.DATA >= " & Formata_Dados(strDataIniAtual, tpDados_DataHora) & _
        " AND ATENDIMENTO.DATA < " & Formata_Dados(strDataFimAtual, tpDados_DataHora) & ")) "
        
    strSql = strSql & " AND PRESTADOR.PKID = " & Formata_Dados(lngPRESTADORID, tpDados_Longo)
    strSql = strSql & " ORDER BY GR.DATA DESC, GR.SEQUENCIAL DESC "
  End If
  '
  Set CapturaGRTurnoCorrentePRE = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGR.CapturaGRTurnoCorrentePRE]", _
            Err.Description
End Function





