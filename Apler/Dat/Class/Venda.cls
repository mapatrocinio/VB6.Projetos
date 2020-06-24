VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirTAB_VENDACARD(lngTAB_VENDACARDID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui estorno, mesmo sabendo que não trabalha com estorno
  strSql = "DELETE FROM TAB_VENDACARD WHERE TAB_VENDACARDEXCID = " & _
    Formata_Dados(lngTAB_VENDACARDID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  'exclui tab_vendacard
  strSql = "DELETE FROM TAB_VENDACARD WHERE PKID = " & _
    Formata_Dados(lngTAB_VENDACARDID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsVenda.ExcluirTAB_VENDACARD]", _
            Err.Description
End Sub

Public Sub UpdateItemTAB_VENDACARD(intITEM As Integer)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE TAB_VENDACARD SET ITEM = ITEM + 1 WHERE ITEM > " & _
    Formata_Dados(intITEM, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsVenda.UpdateItemTAB_VENDACARD]", _
            Err.Description
End Sub

Public Function SelecionarTAB_VENDACARD(lngTAB_VENDACARDID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT * FROM TAB_VENDACARD WHERE PKID = " & _
    Formata_Dados(lngTAB_VENDACARDID, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarTAB_VENDACARD = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsVenda.SelecionarTAB_VENDACARD]", _
            Err.Description
End Function

Public Function CapturaItemTAB_VENDACARD(lngVENDAID As Long) As Long
  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  Dim lngRet As Long
  '
  strSql = "Select MAX(ITEM) As MAXITEM FROM TAB_VENDACARD WHERE VENDAID = " & Formata_Dados(lngVENDAID, tpDados_Longo, tpNulo_NaoAceita) & ";"
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
  CapturaItemTAB_VENDACARD = lngRet
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsVenda.CapturaItemTAB_VENDACARD]", _
            Err.Description
End Function

Public Sub ExlcuirVenda(lngVENDAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM VENDA WHERE PKID = " & _
    Formata_Dados(lngVENDAID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsVenda.ExcluirVenda]", _
            Err.Description
End Sub

Public Function ListarVenda(lngVENDAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT VENDA.*, CONFIGURACAO.TITULO, " & _
    " TOT_PGTO_OUTRA_FORMA = isnull((SELECT COUNT(*) FROM CONTACORRENTE WHERE CONTACORRENTE.VENDAID = VENDA.PKID " & _
    " AND STATUSCC <> 'VD'),0), " & _
    " VR_TOT_VENDA = isnull((SELECT SUM(VALOR) FROM TAB_VENDACARD WHERE TAB_VENDACARD.VENDAID = VENDA.PKID),0) " & _
    " FROM VENDA " & _
    " LEFT JOIN CONFIGURACAO ON CONFIGURACAO.PKID = VENDA.CONFIGURACAOID " & _
    " WHERE VENDA.PKID = " & _
    Formata_Dados(lngVENDAID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarVenda = RunSPReturnRS(strSql)
  If ListarVenda.EOF Then
    ListarVenda.Close
    Set ListarVenda = Nothing
    Err.Raise lngCteErroData, "[" & strClassName & ".clsVenda.ListarVenda]", "A venda não foi gravada."
  End If
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsVenda.ListarVenda]", _
            Err.Description
End Function

Public Sub InserirVenda(lngVENDAID As Long, _
                        strData As String, _
                        strNome As String, _
                        strFuncao As String, _
                        lngTurnoId As Long, _
                        strDescricao As String, _
                        strVenda As String, _
                        strCobranca As String, _
                        strUsuario As String, _
                        lngSEQUENCIAL As Long, _
                        lngCONFIGURACAOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO VENDA (DATA, NOME, FUNCAO, TURNOID, DESCRICAO, VENDA, COBRANCA, USUARIO, SEQUENCIAL, CONFIGURACAOID) VALUES (" & _
    Formata_Dados(strData, tpDados_DataHora, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strFuncao, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngTurnoId, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strVenda, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strCobranca, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strUsuario, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngSEQUENCIAL, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngCONFIGURACAOID, tpDados_Longo) & ");"
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngVENDAID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsVenda.InserirVenda]", _
            Err.Description
End Sub

Public Sub AlterarVenda(lngVENDAID As Long, _
                        strNome As String, _
                        strFuncao As String, _
                        strDescricao As String, _
                        strVenda As String, _
                        strCobranca As String, _
                        lngCONFIGURACAOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE VENDA SET NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "FUNCAO = " & Formata_Dados(strFuncao, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "VENDA = " & Formata_Dados(strVenda, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "COBRANCA = " & Formata_Dados(strCobranca, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "CONFIGURACAOID = " & Formata_Dados(lngCONFIGURACAOID, tpDados_Longo) & " " & _
    "WHERE PKID = " & lngVENDAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsVenda.AlterarVenda]", _
            Err.Description
End Sub


Public Function CapturarVENDAID(lngNumero As Long) As Long
  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  Dim lngRet As Long
  '
  strSql = "Select VENDA.PKID FROM VENDA WHERE SEQUENCIAL = " & Formata_Dados(lngNumero, tpDados_Longo, tpNulo_NaoAceita)
  Set objRs = RunSPReturnRS(strSql)
  If objRs.EOF Then
    objRs.Close
    Set objRs = Nothing
    Err.Raise lngCteErroData, "[" & strClassName & ".clsVenda.CapturarVENDAID]", "Número da Venda não cadastrado."
  Else
    If Not IsNumeric(objRs.Fields("PKID").Value) Then
      objRs.Close
      Set objRs = Nothing
      Err.Raise lngCteErroData, "[" & strClassName & ".clsVenda.CapturarVENDAID]", "Número da Venda não cadastrado."
    Else
      lngRet = objRs.Fields("PKID").Value
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  '
  CapturarVENDAID = lngRet
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsVenda.CapturarVENDAID]", _
            Err.Description
End Function

Public Function ListarTAB_VENDACARD(lngVENDAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT * FROM TAB_VENDACARD WHERE VENDAID = " & _
    Formata_Dados(lngVENDAID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarTAB_VENDACARD = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsVenda.ListarTAB_VENDACARD]", _
            Err.Description
End Function


