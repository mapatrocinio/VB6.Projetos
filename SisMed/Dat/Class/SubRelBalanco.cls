VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRelBalanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Sub SelecionarSaldoBalanco(ByRef curSaldoAnterior As Currency, _
                                  ByRef datDataSaldoAnterior As Date, _
                                  ByRef curReceita As Currency, _
                                  ByRef curPrestador As Currency, _
                                  ByRef curDespesa As Currency, _
                                  ByVal strDataInicial As String, _
                                  ByVal strDataFinal As String)
  On Error GoTo trata
  '
  Dim strSql        As String
  Dim objRs         As ADODB.Recordset
  Dim strDataIni    As String
  Dim strDataFim    As String
  Dim datData       As Date
  'Selecionar Saldo Anterior
  strSql = "SELECT RELBALANCO.* FROM RELBALANCO " & _
      "WHERE DATA = " & _
        "(SELECT MAX(DATA) FROM RELBALANCO " & _
        "WHERE DATA < " & Formata_Dados(strDataInicial, tpDados_DataHora) & ")"
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    'há saldo anteior cadastrado
    curSaldoAnterior = objRs.Fields("SALDO").Value
    datDataSaldoAnterior = objRs.Fields("DATA").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'Obter receita / pago a prestador
  '
  strDataIni = strDataInicial
  datData = CDate(strDataFinal)
  strDataFim = Format(DateAdd("d", 1, datData), "DD/MM/YYYY 00:00")
  
  strSql = "SELECT " & _
    "SUM(CASE vw_cons_t_Financ.ISPGTO " & _
    "WHEN 'S' THEN " & _
    "ISNULL(vw_cons_t_Financ.FINALPRESTCONSCARTAO,0) + ISNULL(vw_cons_t_Financ.FINALPRESTCARTAONAOACEITA,0) " & _
    "    + ISNULL(vw_cons_t_Financ.FINALPRESTCONSESPECIE,0) + ISNULL(vw_cons_t_Financ.FINALPRESTESPECIE,0) + " & _
    "ISNULL(vw_cons_t_Financ.FINALDONOULTRACONSCARTAO,0) + ISNULL(vw_cons_t_Financ.FINALDONOULTRACARTAONAOACEITA,0) " & _
    "    + ISNULL(vw_cons_t_Financ.FINALDONOULTRACONSESPECIE,0) + ISNULL(vw_cons_t_Financ.FINALDONOULTRAESPECIE,0) + " & _
    "ISNULL(vw_cons_t_Financ.FINALDONORXCONSCARTAO,0) + ISNULL(vw_cons_t_Financ.FINALDONORXCARTAONAOACEITA,0) " & _
    "    + ISNULL(vw_cons_t_Financ.FINALDONORXCONSESPECIE,0) + ISNULL(vw_cons_t_Financ.FINALDONORXESPECIE,0) + " & _
    "ISNULL(vw_cons_t_Financ.FINALTECRXCONSCARTAO,0) + ISNULL(vw_cons_t_Financ.FINALTECRXCARTAONAOACEITA,0) " & _
    "    + ISNULL(vw_cons_t_Financ.FINALTECRXCONSESPECIE,0) + ISNULL(vw_cons_t_Financ.FINALTECRXESPECIE,0) " & _
    "ELSE 0 " & _
    "END) AS PGTOPRESTADOR, " & _
    "SUM(CASE GR.STATUS " & _
    "WHEN 'F' THEN " & _
    "ISNULL(vw_cons_t_Financ.PgtoEspecie,0) - ISNULL(vw_cons_t_Financ.PgtoTroco,0) " & _
    "ELSE 0 " & _
    "END) AS PGTODINHEIRO " & _
    "From TURNO " & _
    "INNER JOIN VW_CONS_T_FINANC ON TURNO.PKID = VW_CONS_T_FINANC.TURNOID " & _
    "INNER JOIN GR ON GR.PKID = VW_CONS_T_FINANC.GRID " & _
    "WHERE TURNO.DATA >= " & Formata_Dados(strDataIni, tpDados_DataHora) & _
    "AND TURNO.DATA < " & Formata_Dados(strDataFim, tpDados_DataHora)

  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    'há saldo anteior cadastrado
    curReceita = IIf(IsNull(objRs.Fields("PGTODINHEIRO").Value), 0, objRs.Fields("PGTODINHEIRO").Value)
    curPrestador = IIf(IsNull(objRs.Fields("PGTOPRESTADOR").Value), 0, objRs.Fields("PGTOPRESTADOR").Value)
  End If
  objRs.Close
  Set objRs = Nothing
  'Selecionar Saldo de despesas
  strSql = "SELECT SUM(DESPESA.VR_PAGO) AS VR_PAGO FROM DESPESA " & _
    "WHERE DESPESA.DT_PAGAMENTO >= " & Formata_Dados(strDataIni, tpDados_DataHora) & _
    "AND DESPESA.DT_PAGAMENTO < " & Formata_Dados(strDataFim, tpDados_DataHora)
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    'há saldo anteior cadastrado
    curDespesa = IIf(IsNull(objRs.Fields("VR_PAGO").Value), 0, objRs.Fields("VR_PAGO").Value)
  End If
  objRs.Close
  Set objRs = Nothing
  
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsRelBalanco.SelecionarRelBalanco]"
End Sub

Public Function SelecionarRelBalanco(lngRelBalancoId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT RELBALANCO.* " & _
    " FROM RELBALANCO " & _
    " WHERE RELBALANCO.PKID = " & _
    Formata_Dados(lngRelBalancoId, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarRelBalanco = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsRelBalanco.SelecionarRelBalanco]"
End Function
Public Sub ExcluirRelBalanco(lngRelBalancoId As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM RELBALANCO " & _
    " WHERE PKID = " & lngRelBalancoId
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsRelBalanco.ExcluirRelBalanco]"
End Sub

Public Sub AlterarRelBalanco(lngRelBalancoId As Long, _
                             strData As String, _
                             strSaldo As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE RELBALANCO SET " & _
    " DATA = " & Formata_Dados(strData, tpDados_DataHora) & _
    ", SALDO = " & Formata_Dados(strSaldo, tpDados_Moeda) & _
    " WHERE PKID = " & lngRelBalancoId
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsRelBalanco.AlterarRelBalanco]"
End Sub
Public Sub IncluirRelBalanco(strData As String, _
                             strSaldo As String)
  On Error GoTo trata
  '
  Dim strSql          As String
  '
  strSql = "INSERT INTO RELBALANCO (" & _
    " DATA, SALDO) VALUES (" & _
    Formata_Dados(strData, tpDados_DataHora) & _
    ", " & Formata_Dados(strSaldo, tpDados_Moeda) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsRelBalanco.IncluirRelBalanco]"
End Sub

