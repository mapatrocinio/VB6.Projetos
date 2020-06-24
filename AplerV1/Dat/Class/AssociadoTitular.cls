VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAssociadoTitular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarTitularPeloPkid(lngTitularId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TITULAR.*, ORIGEM.DESCRICAO AS DESCR_ORIGEM, CAPTADOR.NOME AS DESCR_CAPTADOR, EMPRESA.NOME AS DESCR_EMPRESA  " & _
    "FROM TITULAR LEFT JOIN ORIGEM ON TITULAR.ORIGEMID = ORIGEM.PKID " & _
    "LEFT JOIN CAPTADOR ON TITULAR.CAPTADORID = CAPTADOR.PKID " & _
        "LEFT JOIN EMPRESA ON TITULAR.EMPRESAID = EMPRESA.PKID " & _
    "WHERE TITULAR.ASSOCIADOID = " & _
    Formata_Dados(lngTitularId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarTitularPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTitular.SelecionarTitularPeloPkid]", _
            Err.Description
End Function

Public Sub AlterarTitular(lngPKID As Long, _
                          lngORIGEMID As Long, _
                          lngCaptadorId As Long, _
                          lngEMPRESAID As Long, _
                          strMatricula As String, _
                          strNumeroContrato As String, _
                          strDataIniContrato As String, _
                          strDataFimContrato As String, _
                          strEndRuaCom As String, _
                          strEndNumeroCom As String, _
                          strEndComplCom As String, _
                          strEndCepCom As String, _
                          strEndBairroCom As String, _
                          strEndCidadeCom As String, _
                          strEndEstadoCom As String, _
                          strTelefoneCom1 As String, _
                          strTelefoneCom2 As String, _
                          strEndRuaCob As String, _
                          strEndNumeroCob As String, _
                          strEndComplCob As String, _
                          strEndCepCob As String, _
                          strEndBairroCob As String, _
                          strEndCidadeCob As String, _
                          strEndEstadoCob As String)


  On Error GoTo trata
  '
  Dim strSql As String
  '
  strSql = "UPDATE TITULAR SET " & _
    "ORIGEMID = " & Formata_Dados(IIf(lngORIGEMID = 0, "", lngORIGEMID), tpDados_Longo) & ", " & _
    "CAPTADORID = " & Formata_Dados(IIf(lngCaptadorId = 0, "", lngCaptadorId), tpDados_Longo) & ", " & _
    "EMPRESAID = " & Formata_Dados(IIf(lngEMPRESAID = 0, "", lngEMPRESAID), tpDados_Longo) & ", " & _
    "MATRICULA = " & Formata_Dados(strMatricula, tpDados_Moeda) & ", " & _
    "NUMEROCONTRATO = " & Formata_Dados(strNumeroContrato, tpDados_Texto) & ", " & _
    "DATAINICONTRATO = " & Formata_Dados(strDataIniContrato, tpDados_DataHora) & ", " & _
    "DATAFIMCONTRATO = " & Formata_Dados(strDataFimContrato, tpDados_DataHora) & ", " & _
    "ENDRUACOM = " & Formata_Dados(strEndRuaCom, tpDados_Texto) & ", " & _
    "ENDNUMEROCOM = " & Formata_Dados(strEndNumeroCom, tpDados_Texto) & ", " & _
    "ENDCOMPLCOM = " & Formata_Dados(strEndComplCom, tpDados_Texto) & ", " & _
    "ENDCEPCOM = " & Formata_Dados(strEndCepCom, tpDados_Texto) & ", " & _
    "ENDBAIRROCOM = " & Formata_Dados(strEndBairroCom, tpDados_Texto) & ", " & _
    "ENDCIDADECOM = " & Formata_Dados(strEndCidadeCom, tpDados_Texto) & ", " & _
    "ENDESTADOCOM = " & Formata_Dados(strEndEstadoCom, tpDados_Texto) & ", " & _
    "TELEFONECOM1 = " & Formata_Dados(strTelefoneCom1, tpDados_Texto) & ", " & _
    "TELEFONECOM2 = " & Formata_Dados(strTelefoneCom2, tpDados_Texto) & ", "
  strSql = strSql & "ENDRUACOB = " & Formata_Dados(strEndRuaCob, tpDados_Texto) & ", " & _
    "ENDNUMEROCOB = " & Formata_Dados(strEndNumeroCob, tpDados_Texto) & ", " & _
    "ENDCOMPLCOB = " & Formata_Dados(strEndComplCob, tpDados_Texto) & ", " & _
    "ENDCEPCOB = " & Formata_Dados(strEndCepCob, tpDados_Texto) & ", " & _
    "ENDBAIRROCOB = " & Formata_Dados(strEndBairroCob, tpDados_Texto) & ", " & _
    "ENDCIDADECOB = " & Formata_Dados(strEndCidadeCob, tpDados_Texto) & ", " & _
    "ENDESTADOCOB = " & Formata_Dados(strEndEstadoCob, tpDados_Texto) & " " & _
    "WHERE ASSOCIADOID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTitular.AlterarTitular]", _
            Err.Description
End Sub

Public Sub InserirTitular(lngPKID, _
                          lngORIGEMID As Long, _
                          lngCaptadorId As Long, _
                          lngEMPRESAID As Long, _
                          strMatricula As String, _
                          strNumeroContrato As String, _
                          strDataIniContrato As String, _
                          strDataFimContrato As String, _
                          strEndRuaCom As String, _
                          strEndNumeroCom As String, _
                          strEndComplCom As String, _
                          strEndCepCom As String, _
                          strEndBairroCom As String, _
                          strEndCidadeCom As String, _
                          strEndEstadoCom As String, _
                          strTelefoneCom1 As String, _
                          strTelefoneCom2 As String, _
                          strEndRuaCob As String, _
                          strEndNumeroCob As String, _
                          strEndComplCob As String, _
                          strEndCepCob As String, _
                          strEndBairroCob As String, _
                          strEndCidadeCob As String, _
                          strEndEstadoCob As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO TITULAR (ASSOCIADOID, ORIGEMID, CAPTADORID, EMPRESAID, MATRICULA, NUMEROCONTRATO, DATAINICONTRATO, DATAFIMCONTRATO, ENDRUACOM, ENDNUMEROCOM, ENDCOMPLCOM, ENDCEPCOM, ENDBAIRROCOM, ENDCIDADECOM, " & _
    "ENDESTADOCOM, TELEFONECOM1, TELEFONECOM2, ENDRUACOB, ENDNUMEROCOB, ENDCOMPLCOB, ENDCEPCOB, ENDBAIRROCOB, ENDCIDADECOB, ENDESTADOCOB) VALUES(" & _
    Formata_Dados(lngPKID, tpDados_Longo) & ", " & _
    Formata_Dados(IIf(lngORIGEMID = 0, "", lngORIGEMID), tpDados_Longo) & ", " & _
    Formata_Dados(IIf(lngCaptadorId = 0, "", lngCaptadorId), tpDados_Longo) & ", " & _
    Formata_Dados(IIf(lngEMPRESAID = 0, "", lngEMPRESAID), tpDados_Longo) & ", " & _
    Formata_Dados(strMatricula, tpDados_Moeda) & ", " & _
    Formata_Dados(strNumeroContrato, tpDados_Texto) & ", " & _
    Formata_Dados(strDataIniContrato, tpDados_DataHora) & ", " & _
    Formata_Dados(strDataFimContrato, tpDados_DataHora) & ", " & _
    Formata_Dados(strEndRuaCom, tpDados_Texto) & ", " & _
    Formata_Dados(strEndNumeroCom, tpDados_Texto) & ", " & _
    Formata_Dados(strEndComplCom, tpDados_Texto) & ", " & _
    Formata_Dados(strEndCepCom, tpDados_Texto) & ", " & _
    Formata_Dados(strEndBairroCom, tpDados_Texto) & ", " & _
    Formata_Dados(strEndCidadeCom, tpDados_Texto) & ", " & _
    Formata_Dados(strEndEstadoCom, tpDados_Texto) & ", " & _
    Formata_Dados(strTelefoneCom1, tpDados_Texto) & ", " & _
    Formata_Dados(strTelefoneCom2, tpDados_Texto) & ", "
  strSql = strSql & Formata_Dados(strEndRuaCob, tpDados_Texto) & ", " & _
    Formata_Dados(strEndNumeroCob, tpDados_Texto) & ", " & _
    Formata_Dados(strEndComplCob, tpDados_Texto) & ", " & _
    Formata_Dados(strEndCepCob, tpDados_Texto) & ", " & _
    Formata_Dados(strEndBairroCob, tpDados_Texto) & ", " & _
    Formata_Dados(strEndCidadeCob, tpDados_Texto) & ", " & _
    Formata_Dados(strEndEstadoCob, tpDados_Texto) & ")"
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTitular.InserirTitular]", _
            Err.Description
End Sub

Public Function SelecionarTitular(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TITULAR.*, ORIGEM.DESCRICAO AS DESCR_ORIGEM, CAPTADOR.NOME AS DESCR_CAPTADOR " & _
    "FROM TITULAR LEFT JOIN ORIGEM ON TITULAR.ORIGEMID = ORIGEM.PKID " & _
    "LEFT JOIN CAPTADOR ON TITULAR.CAPTADORID = CAPTADOR.PKID " & _
    "WHERE TITULAR.ASSOCIADOID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarTitular = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTitular.SelecionarTitular]", _
            Err.Description
End Function

