VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExtraForaUnidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirTAB_EXTRAUNIDADE(ByVal lngEXTRAUNIDADEID As Long, _
                                   ByVal strQuantidade As String, _
                                   ByVal strValor As String, _
                                   ByVal strItem As String, _
                                   ByVal strDescricao As String, _
                                   ByVal strTAB_EXTRAUNIDADEEXCID As String)
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "INSERT INTO TAB_EXTRAUNIDADE (EXTRAUNIDADEID, QUANTIDADE, VALOR, ITEM, DESCRICAO, TAB_EXTRAUNIDADEEXCID) VALUES (" & _
    Formata_Dados(lngEXTRAUNIDADEID, tpDados_Longo, tpNulo_Aceita) & ", " & _
    Formata_Dados(strQuantidade, tpDados_Longo, tpNulo_Aceita) & ", " & _
    Formata_Dados(strValor, tpDados_Moeda, tpNulo_Aceita) & ", " & _
    Formata_Dados(strItem, tpDados_Longo, tpNulo_Aceita) & ", " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strTAB_EXTRAUNIDADEEXCID, tpDados_Longo, tpNulo_Aceita) & ");"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsExtraForaUnidade.InserirTAB_EXTRAUNIDADE]", _
            Err.Description
End Sub

Public Sub ExcluirTAB_EXTRAUNIDADE(lngTAB_EXTRAUNIDADEID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui estorno, mesmo sabendo que não trabalha com estorno
  strSql = "DELETE FROM TAB_EXTRAUNIDADE WHERE TAB_EXTRAUNIDADEEXCID = " & _
    Formata_Dados(lngTAB_EXTRAUNIDADEID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  'exclui tab_EXTRAUNIDADE
  strSql = "DELETE FROM TAB_EXTRAUNIDADE WHERE PKID = " & _
    Formata_Dados(lngTAB_EXTRAUNIDADEID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsExtraForaUnidade.ExcluirTAB_EXTRAUNIDADE]", _
            Err.Description
End Sub

Public Sub UpdateItemTAB_EXTRAUNIDADE(intITEM As Integer)
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE TAB_EXTRAUNIDADE SET ITEM = ITEM + 1 WHERE ITEM > " & _
    Formata_Dados(intITEM, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsExtraForaUnidade.UpdateItemTAB_EXTRAUNIDADE]", _
            Err.Description
End Sub

Public Function SelecionarTAB_EXTRAUNIDADE(lngTAB_EXTRAUNIDADEID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT * FROM TAB_EXTRAUNIDADE WHERE PKID = " & _
    Formata_Dados(lngTAB_EXTRAUNIDADEID, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarTAB_EXTRAUNIDADE = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsExtraForaUnidade.SelecionarTAB_EXTRAUNIDADE]", _
            Err.Description
End Function

Public Function CapturaItemTAB_EXTRAUNIDADE(lngEXTRAUNIDADEID As Long) As Long
  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  Dim lngRet As Long
  '
  strSql = "Select MAX(ITEM) As MAXITEM FROM TAB_EXTRAUNIDADE WHERE EXTRAUNIDADEID = " & Formata_Dados(lngEXTRAUNIDADEID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  Set objRs = RunSPReturnRS(strSql)
  If objRs.EOF Then
    lngRet = 1
  Else
    If Not IsNumeric(objRs.Fields("MAXITEM").Value) Then
      lngRet = 1
    Else
      lngRet = objRs.Fields("MAXITEM").Value + 1
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  '
  CapturaItemTAB_EXTRAUNIDADE = lngRet
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsExtraForaUnidade.CapturaItemTAB_EXTRAUNIDADE]", _
            Err.Description
End Function

Public Sub ExlcuirExtra(lngEXTRAUNIDADEID As Long)
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "DELETE FROM EXTRAUNIDADE WHERE PKID = " & _
    Formata_Dados(lngEXTRAUNIDADEID, tpDados_Longo, tpNulo_NaoAceita) & ";"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsExtraForaUnidade.ExcluirExtra]", _
            Err.Description
End Sub

Public Function ListarExtra(lngEXTRAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT EXTRAUNIDADE.*, CONFIGURACAO.TITULO, " & _
    " TOT_PGTO_OUTRA_FORMA = isnull((SELECT COUNT(*) FROM CONTACORRENTE WHERE CONTACORRENTE.EXTRAID = EXTRAUNIDADE.PKID " & _
    " AND STATUSCC <> 'EX'),0), " & _
    " VR_TOT_EXTRA = isnull((SELECT SUM(VALOR) FROM TAB_EXTRAUNIDADE WHERE TAB_EXTRAUNIDADE.EXTRAUNIDADEID = EXTRAUNIDADE.PKID),0) " & _
    " FROM EXTRAUNIDADE " & _
    " INNER JOIN CONFIGURACAO ON CONFIGURACAO.PKID = EXTRAUNIDADE.CONFIGURACAOID " & _
    " WHERE EXTRAUNIDADE.PKID = " & _
    Formata_Dados(lngEXTRAID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarExtra = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsExtra.ListarExtra]", _
            Err.Description
End Function

Public Sub InserirExtra(strData As String, _
                        lngTurnoId As Long, _
                        strDescricao As String, _
                        strTipo As String, _
                        lngSEQUENCIAL As Long, _
                        lngCONFIGURACAOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "INSERT INTO EXTRAUNIDADE (DTEXTRA, TURNOID, DESCRICAO, TIPO, NRO, CONFIGURACAOID) VALUES (" & _
    Formata_Dados(strData, tpDados_DataHora, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngTurnoId, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strTipo, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngSEQUENCIAL, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngCONFIGURACAOID, tpDados_Longo) & _
    ");"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsExtraForaunidade.InserirExtra]", _
            Err.Description
End Sub

Public Sub AlterarExtra(lngEXTRAID As Long, _
                        strDescricao As String, _
                        lngCONFIGURACAOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE EXTRAUNIDADE SET " & _
    "DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "CONFIGURACAOID = " & Formata_Dados(lngCONFIGURACAOID, tpDados_Longo) & _
    " WHERE PKID = " & lngEXTRAID

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsExtraForaUnidade.AlterarExtra]", _
            Err.Description
End Sub


Public Function CapturarEXTRAID(lngNumero As Long) As Long
  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "Select MAX(EXTRAUNIDADE.PKID) AS PKID FROM EXTRAUNIDADE WHERE NRO = " & Formata_Dados(lngNumero, tpDados_Longo, tpNulo_NaoAceita)
  Set objRs = RunSPReturnRS(strSql)
  If objRs.EOF Then
    CapturarEXTRAID = 0
  Else
    If Not IsNumeric(objRs.Fields("PKID").Value) Then
      CapturarEXTRAID = 0
    Else
      CapturarEXTRAID = objRs.Fields("PKID").Value
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsExtraForaUnidade.CapturarEXTRAID]", _
            Err.Description
End Function

Public Function ListarTAB_EXTRAUNIDADE(lngEXTRAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT * FROM TAB_EXTRAUNIDADE WHERE EXTRAUNIDADEID = " & _
    Formata_Dados(lngEXTRAID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarTAB_EXTRAUNIDADE = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsExtraForaUnidade.ListarTAB_EXTRAUNIDADE]", _
            Err.Description
End Function


