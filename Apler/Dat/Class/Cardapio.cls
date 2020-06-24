VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCardapio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Sub ExcluirTAB_CARDESTINTER(lngTAB_CARDESTINTERID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui TAB_CARDESTINTER
  strSql = "DELETE FROM TAB_CARDESTINTER WHERE PKID = " & _
    Formata_Dados(lngTAB_CARDESTINTERID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.ExcluirTAB_CARDESTINTER]", _
            Err.Description
End Sub

Public Sub InserirTAB_CARDESTINTER(ByVal lngCARDAPIORESUMOID As Long, _
                                   ByVal lngESTOQUEINTERMEDIARIOID As Long, _
                                   ByVal strQuantidade As String, _
                                   ByVal strTipo As String)
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "INSERT INTO TAB_CARDESTINTER (CARDAPIO_RESUMOID, ESTOQUEINTERMEDIARIOID, QUANTIDADE, TIPO) VALUES (" & _
    Formata_Dados(lngCARDAPIORESUMOID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngESTOQUEINTERMEDIARIOID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strQuantidade, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strTipo, tpDados_Texto, tpNulo_Aceita) & ");"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.InserirTAB_CARDESTINTER]", _
            Err.Description
End Sub


Public Function CapturaEstoqueIntermediarioPeloCodigo(ByVal strCodigo As String) As ADODB.Recordset
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "SELECT * FROM ESTOQUEINTERMEDIARIO WHERE CODIGO = " & Formata_Dados(strCodigo, tpDados_Texto, tpNulo_NaoAceita)
  Set CapturaEstoqueIntermediarioPeloCodigo = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.CapturaEstoqueIntermediarioPeloCodigo]", _
            Err.Description
End Function
Public Sub InserirCardapio(strCodigo As String, _
                           strDescricao As String, _
                           strValor As String, _
                           strGrupo As String, _
                           blnPreservativoLei As Boolean, _
                           strTipo As String, _
                           strTipoDiretoria As String, _
                           strAliquota As String, _
                           blnExcluido As Boolean, _
                           blnCortesia As Boolean, _
                           strData As String, _
                           lngCARDAPIORESUMOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO CARDAPIO (CODIGO, DESCRICAO, VALOR, DATA, GRUPO, PRESERVATIVOLEI, " & _
    "TIPO, TIPO_DIRETORIA, ALIQUOTA, EXCLUIDO, CORTESIA, CARDAPIO_RESUMOID) VALUES (" & _
    Formata_Dados(strCodigo, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strValor, tpDados_Moeda, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strData, tpDados_DataHora, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strGrupo, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(blnPreservativoLei, tpDados_Boolean, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strTipo, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strTipoDiretoria, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strAliquota, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(blnExcluido, tpDados_Boolean, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(blnCortesia, tpDados_Boolean, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngCARDAPIORESUMOID, tpDados_Longo, tpNulo_NaoAceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.InserirCardapio]", _
            Err.Description
End Sub

Public Sub AlterarCardapio(lngCARDAPIOID As Long, _
                           strCodigo As String, _
                           strDescricao As String, _
                           strValor As String, _
                           strGrupo As String, _
                           blnPreservativoLei As Boolean, _
                           strTipo As String, _
                           strTipoDiretoria As String, _
                           strAliquota As String, _
                           blnExcluido As Boolean, _
                           blnCortesia As Boolean)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE CARDAPIO SET CODIGO = " & Formata_Dados(strCodigo, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "VALOR = " & Formata_Dados(strValor, tpDados_Moeda, tpNulo_NaoAceita) & ", " & _
    "GRUPO = " & Formata_Dados(strGrupo, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "PRESERVATIVOLEI = " & Formata_Dados(blnPreservativoLei, tpDados_Boolean, tpNulo_NaoAceita) & ", " & _
    "TIPO = " & Formata_Dados(strTipo, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "TIPO_DIRETORIA = " & Formata_Dados(strTipoDiretoria, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "ALIQUOTA = " & Formata_Dados(strAliquota, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "EXCLUIDO = " & Formata_Dados(blnExcluido, tpDados_Boolean, tpNulo_NaoAceita) & ", " & _
    "CORTESIA = " & Formata_Dados(blnCortesia, tpDados_Boolean, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngCARDAPIOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.AlterarCardapio]", _
            Err.Description
End Sub

Public Function ListarTAB_CARDESTINTER(lngCARDAPIOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  'strSql = "SELECT TAB_CARDESTINTER.*, TIPO.DESCRICAO AS DESCTIPO " & _
    "FROM ((CARDAPIO LEFT JOIN CARDAPIO_RESUMO ON CARDAPIO.CARDAPIO_RESUMOID = CARDAPIO_RESUMO.PKID) " & _
    "LEFT JOIN TAB_CARDESTINTER ON CARDAPIO_RESUMO.PKID = TAB_CARDESTINTER.CARDAPIO_RESUMOID) " & _
    "LEFT JOIN TIPO ON TIPO.TIPO = TAB_CARDESTINTER.TIPO " & _
    "WHERE CARDAPIO.PKID = " & _
    Formata_Dados(lngCARDAPIOID, tpDados_Longo, tpNulo_NaoAceita)
  'strSql = "SELECT TAB_CARDESTINTER.*, TIPO.DESCRICAO AS DESCTIPO " & _
    "FROM ((CARDAPIO INNER JOIN CARDAPIO_RESUMO ON CARDAPIO.CARDAPIO_RESUMOID = CARDAPIO_RESUMO.PKID) " & _
    "INNER JOIN TAB_CARDESTINTER ON CARDAPIO_RESUMO.PKID = TAB_CARDESTINTER.CARDAPIO_RESUMOID) " & _
    "INNER JOIN TIPO ON TIPO.TIPO = TAB_CARDESTINTER.TIPO " & _
    "WHERE CARDAPIO.PKID = " & _
    Formata_Dados(lngCARDAPIOID, tpDados_Longo, tpNulo_NaoAceita)
  strSql = "SELECT TAB_CARDESTINTER.*, TIPO.DESCRICAO AS DESCTIPO " & _
    "FROM ((CARDAPIO INNER JOIN CARDAPIO_RESUMO ON CARDAPIO.CARDAPIO_RESUMOID = CARDAPIO_RESUMO.PKID) " & _
    "INNER JOIN TAB_CARDESTINTER ON CARDAPIO_RESUMO.PKID = TAB_CARDESTINTER.CARDAPIO_RESUMOID) " & _
    "LEFT JOIN TIPO ON TIPO.TIPO = TAB_CARDESTINTER.TIPO " & _
    "WHERE CARDAPIO.PKID = " & _
    Formata_Dados(lngCARDAPIOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarTAB_CARDESTINTER = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.ListarTAB_CARDESTINTER]", _
            Err.Description
End Function

Public Function ListarCardapioPeloTipo(lngCARDAPIORESUMOID As Long, _
                                       strTipo As String, _
                                       strTipoDiretoria As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CARDAPIO.* " & _
    "FROM CARDAPIO " & _
    "WHERE CARDAPIO.CARDAPIO_RESUMOID = " & _
    Formata_Dados(lngCARDAPIORESUMOID, tpDados_Longo, tpNulo_NaoAceita) & _
    " AND CARDAPIO.TIPO = " & Formata_Dados(strTipo, tpDados_Texto, tpNulo_NaoAceita) & _
    " AND CARDAPIO.TIPO_DIRETORIA = " & Formata_Dados(strTipoDiretoria, tpDados_Texto, tpNulo_NaoAceita)

  Set ListarCardapioPeloTipo = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.ListarCardapioPeloTipo]", _
            Err.Description
End Function

Public Function ListarCardapio(lngCARDAPIOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CARDAPIO.*, TIPO.DESCRICAO AS DESCTIPO, TIPO_DIRETORIA.DESCRICAO AS DESCTIPODIR " & _
    "FROM (CARDAPIO LEFT JOIN TIPO ON TIPO.TIPO = CARDAPIO.TIPO) " & _
    "LEFT JOIN TIPO_DIRETORIA ON TIPO_DIRETORIA.TIPO_DIRETORIA = CARDAPIO.TIPO_DIRETORIA " & _
    "WHERE CARDAPIO.PKID = " & _
    Formata_Dados(lngCARDAPIOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarCardapio = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.ListarCardapio]", _
            Err.Description
End Function

Public Sub ExcluirCARDAPIO(lngCARDAPIOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui CARDÁPIO
  strSql = "DELETE FROM CARDAPIO WHERE CARDAPIO.PKID = " & _
    Formata_Dados(lngCARDAPIOID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.ExcluirCARDAPIO]", _
            Err.Description
End Sub



Public Sub InserirTipoDiretoria(strTipoDiretoria As String, _
                                strDescricao As String, _
                                strAliquotaPadrao As String, _
                                strPercentualPadrao As String)


  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO TIPO_DIRETORIA (TIPO_DIRETORIA, DESCRICAO, ALIQUOTAPADRAO, PERCENTUALPADRAO) VALUES (" & _
    Formata_Dados(strTipoDiretoria, tpDados_Texto, tpNulo_NaoAceita) & _
    ", " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & _
    ", " & Formata_Dados(strAliquotaPadrao, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strPercentualPadrao, tpDados_Longo, tpNulo_Aceita) & " )"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.InserirTipoDiretoria]", _
            Err.Description
End Sub

Public Sub InserirTipo(strTipo As String, _
                       strDescricao As String, _
                       strTipoCardapio As String, _
                       strIndTipo As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO TIPO (TIPO, DESCRICAO, TIPO_CARDAPIO, INDTIPO) VALUES (" & _
    Formata_Dados(strTipo, tpDados_Texto, tpNulo_NaoAceita) & _
    ", " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & _
    ", " & Formata_Dados(strTipoCardapio, tpDados_Texto, tpNulo_NaoAceita) & _
    ", " & Formata_Dados(strIndTipo, tpDados_Texto) & " )"
    
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.InserirTipo]", _
            Err.Description
End Sub

Public Sub InserirGrupoCardapio(strNome As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO GRUPOCARDAPIO (NOME) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.InserirGrupoCardapio]", _
            Err.Description
End Sub

Public Sub AlterarGrupoCardapio(lngGrupoCardapioId As Long, _
                                strNome As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE GRUPOCARDAPIO SET NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngGrupoCardapioId
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.AlterarGrupoCardapio]", _
            Err.Description
End Sub

                                
Public Sub AlterarTipoDiretoria(strTipoDiretoria As String, _
                                strDescricao As String, _
                                strAliquotaPadrao As String, _
                                strPercentualPadrao As String, _
                                strTipoAnterior As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE TIPO_DIRETORIA SET TIPO_DIRETORIA = " & Formata_Dados(strTipoDiretoria, tpDados_Texto, tpNulo_NaoAceita) & _
    ", DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & _
    ", ALIQUOTAPADRAO = " & Formata_Dados(strAliquotaPadrao, tpDados_Texto, tpNulo_Aceita) & _
    ", PERCENTUALPADRAO = " & Formata_Dados(strPercentualPadrao, tpDados_Longo, tpNulo_Aceita) & _
    " WHERE TIPO_DIRETORIA = " & Formata_Dados(strTipoAnterior, tpDados_Texto, tpNulo_NaoAceita)
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.AlterarTipoDiretoria]", _
            Err.Description
End Sub

Public Sub AlterarTipo(strTipo As String, _
                       strDescricao As String, _
                       strTipoCardapio As String, _
                       strTipoAnterior As String, _
                       strIndTipo As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE TIPO SET TIPO = " & Formata_Dados(strTipo, tpDados_Texto, tpNulo_NaoAceita) & _
    ", DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & _
    ", TIPO_CARDAPIO = " & Formata_Dados(strTipoCardapio, tpDados_Texto, tpNulo_NaoAceita) & _
    ", INDTIPO = " & Formata_Dados(strIndTipo, tpDados_Texto) & _
    " WHERE TIPO = " & Formata_Dados(strTipoAnterior, tpDados_Texto, tpNulo_NaoAceita)
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.AlterarTipo]", _
            Err.Description
End Sub

Public Function ListarGrupoCardapio(lngGrupoCardapioId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT GRUPOCARDAPIO.* " & _
    "FROM GRUPOCARDAPIO " & _
    "WHERE GRUPOCARDAPIO.PKID = " & _
    Formata_Dados(lngGrupoCardapioId, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarGrupoCardapio = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.ListarGrupoCardapio]", _
            Err.Description
End Function

Public Function ListarTipoDiretoriaPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TIPO_DIRETORIA.* " & _
    "FROM TIPO_DIRETORIA " & _
    "WHERE TIPO_DIRETORIA.DESCRICAO = " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarTipoDiretoriaPelaDesc = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.ListarTipoDiretoriaPelaDesc]", _
            Err.Description
End Function

Public Function ListarTipoPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TIPO.* " & _
    "FROM TIPO " & _
    "WHERE TIPO.DESCRICAO = " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarTipoPelaDesc = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.ListarTipoPelaDesc]", _
            Err.Description
End Function

Public Function ListarTipo(strTipo As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TIPO.* " & _
    "FROM TIPO " & _
    "WHERE TIPO.TIPO = " & _
    Formata_Dados(strTipo, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarTipo = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.ListarTipo]", _
            Err.Description
End Function

Public Function ListarTipoDiretoria(strTipoDiretoria As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT TIPO_DIRETORIA.* " & _
    "FROM TIPO_DIRETORIA " & _
    "WHERE TIPO_DIRETORIA.TIPO_DIRETORIA = " & _
    Formata_Dados(strTipoDiretoria, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarTipoDiretoria = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.ListarTipoDiretoria]", _
            Err.Description
End Function


Public Sub ExcluirGRUPOCARDAPIO(lngGrupoCardapioId As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui GRUPO CARDÁPIO
  strSql = "DELETE FROM GRUPOCARDAPIO WHERE GRUPOCARDAPIO.PKID = " & _
    Formata_Dados(lngGrupoCardapioId, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.ExcluirGRUPOCARDAPIO]", _
            Err.Description
End Sub

Public Sub ExcluirTipo(strTipo As String)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui Tipo de estoque
  strSql = "DELETE FROM TIPO WHERE TIPO.TIPO = " & _
    Formata_Dados(strTipo, tpDados_Texto, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.ExcluirTipo]", _
            Err.Description
End Sub

Public Sub ExcluirTipoDiretoria(strTipoDiretoria As String)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui Tipo de estoque
  strSql = "DELETE FROM TIPO_DIRETORIA WHERE TIPO_DIRETORIA.TIPO_DIRETORIA = " & _
    Formata_Dados(strTipoDiretoria, tpDados_Texto, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.ExcluirTipoDiretoria]", _
            Err.Description
End Sub

Public Sub InserirTAB_VENDACARD(ByVal lngVENDAID As Long, _
                                ByVal lngCARDAPIOID As Long, _
                                ByVal strQuantidade As String, _
                                ByVal strValor As String, _
                                ByVal strItem As String, _
                                ByVal strTAB_VENDACARDEXCID As String, _
                                ByVal strDescricao As String)
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "INSERT INTO TAB_VENDACARD (VENDAID, CARDAPIOID, QUANTIDADE, VALOR, ITEM, DESCRICAO, TAB_VENDACARDEXCID) VALUES (" & _
    Formata_Dados(lngVENDAID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngCARDAPIOID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strQuantidade, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strValor, tpDados_Moeda, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strItem, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strTAB_VENDACARDEXCID, tpDados_Longo, tpNulo_Aceita) & ");"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.InserirTAB_VENDACARD]", _
            Err.Description
End Sub

Public Function CapturaCortesiaCardapioVda(ByVal lngVENDAID As Long) As ADODB.Recordset
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "Select * From (VENDA INNER JOIN TAB_VENDACARD ON VENDA.PKID = TAB_VENDACARD.VENDAID) INNER JOIN CARDAPIO ON CARDAPIO.PKID = TAB_VENDACARD.CARDAPIOID " & _
                  " WHERE VENDA.PKID = " & Formata_Dados(lngVENDAID, tpDados_Longo, tpNulo_NaoAceita) & " AND CARDAPIO.PreservativoLei = False AND CARDAPIO!Tipo_Diretoria = 'C'"
  Set CapturaCortesiaCardapioVda = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.CapturaCortesiaCardapioVda]", _
            Err.Description
End Function


Public Sub InserirTAB_PEDIDOCARD(ByVal lngPEDIDOID As Long, _
                                 ByVal lngCARDAPIOID As Long, _
                                 ByVal strQuantidade As String, _
                                 ByVal strValor As String, _
                                 ByVal strItem As String, _
                                 ByVal strTAB_PEDIDOCARDEXCID As String, _
                                 ByVal strDescricao As String)
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "INSERT INTO TAB_PEDIDOCARD (PEDIDOID, CARDAPIOID, QUANTIDADE, VALOR, ITEM, DESCRICAO, TAB_PEDIDOCARDEXCID) VALUES (" & _
    Formata_Dados(lngPEDIDOID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngCARDAPIOID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strQuantidade, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strValor, tpDados_Moeda, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strItem, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strTAB_PEDIDOCARDEXCID, tpDados_Longo, tpNulo_Aceita) & ");"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.InserirTAB_PEDIDOCARD]", _
            Err.Description
End Sub
Public Function CapturaCortesiaCardapio(ByVal lngLocacaoId As Long) As ADODB.Recordset
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "Select * From (PEDIDO INNER JOIN TAB_PEDIDOCARD ON PEDIDO.PKID = TAB_PEDIDOCARD.PEDIDOID) INNER JOIN CARDAPIO ON CARDAPIO.PKID = TAB_PEDIDOCARD.CARDAPIOID " & _
                  " WHERE PEDIDO.ALOCACAOID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_NaoAceita) & _
                  " AND CARDAPIO.PreservativoLei = " & Formata_Dados(False, tpDados_Boolean) & _
                  " AND CARDAPIO.Tipo_Diretoria = " & Formata_Dados("C", tpDados_Texto)
  Set CapturaCortesiaCardapio = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.CapturaCortesiaCardapio]", _
            Err.Description
End Function

Public Function CapturaItemCardapio(ByVal strCodigo As String) As ADODB.Recordset
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "SELECT * FROM CARDAPIO WHERE CODIGO LIKE " & Formata_Dados("%" & strCodigo & "%", tpDados_Texto, tpNulo_NaoAceita) & _
    " OR DESCRICAO LIKE " & Formata_Dados("%" & strCodigo & "%", tpDados_Texto, tpNulo_NaoAceita)
  
  Set CapturaItemCardapio = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.CapturaItemCardapio]", _
            Err.Description
End Function

Public Function CapturaQtdPreservativoLeidoCardapio(ByVal lngLocacaoId As Long) As ADODB.Recordset
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "Select Sum(TAB_PEDIDOCARD.Quantidade) AS QtdPreservativoLei From (PEDIDO INNER JOIN TAB_PEDIDOCARD ON PEDIDO.PKID = TAB_PEDIDOCARD.PEDIDOID) INNER JOIN CARDAPIO ON CARDAPIO.PKID = TAB_PEDIDOCARD.CARDAPIOID " & _
          " WHERE PEDIDO.ALOCACAOID = " & Formata_Dados(lngLocacaoId, tpDados_Longo, tpNulo_NaoAceita) & " AND CARDAPIO.PreservativoLei = True"
  Set CapturaQtdPreservativoLeidoCardapio = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.CapturaQtdPreservativoLeidoCardapio]", _
            Err.Description
End Function

Public Function CapturaTab_CardEstInterCardapio(ByVal lngCARDAPIOID As Long, _
                                                ByVal lngTAB_CARDESTINTERID As Long, _
                                                ByVal lngGRUPOESTOQUEID As Long) As ADODB.Recordset
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "Select TAB_GRUPOESTESTINTER.PKID AS TAB_GRUPOESTESTINTERID, ESTOQUEINTERMEDIARIO.*, TAB_CARDESTINTER.QUANTIDADE, TAB_GRUPOESTESTINTER.QTDESTOQUE " & _
    "From (((ESTOQUEINTERMEDIARIO INNER JOIN TAB_CARDESTINTER ON ESTOQUEINTERMEDIARIO.PKID = TAB_CARDESTINTER.ESTOQUEINTERMEDIARIOID) " & _
    " INNER JOIN TAB_GRUPOESTESTINTER ON ESTOQUEINTERMEDIARIO.PKID = TAB_GRUPOESTESTINTER.ESTOQUEINTERMEDIARIOID) " & _
    " INNER JOIN CARDAPIO_RESUMO ON CARDAPIO_RESUMO.PKID = TAB_CARDESTINTER.CARDAPIO_RESUMOID) " & _
    " INNER JOIN CARDAPIO ON CARDAPIO_RESUMO.PKID = CARDAPIO.CARDAPIO_RESUMOID " & _
    " WHERE CARDAPIO.PKID = " & Formata_Dados(lngCARDAPIOID, tpDados_Longo, tpNulo_NaoAceita) & _
    " AND TAB_GRUPOESTESTINTER.GRUPOESTOQUEID = " & Formata_Dados(lngGRUPOESTOQUEID, tpDados_Longo, tpNulo_NaoAceita) & _
    " AND TAB_CARDESTINTER.PKID = " & Formata_Dados(lngTAB_CARDESTINTERID, tpDados_Longo, tpNulo_NaoAceita)
  '
  Set CapturaTab_CardEstInterCardapio = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.CapturaTab_CardEstInterCardapio]", _
            Err.Description
End Function

Public Function VerificaProdutoCardapioEst(ByVal strCodigo As String) As ADODB.Recordset
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "SELECT * FROM CARDAPIO WHERE CODIGO = " & Formata_Dados(strCodigo, tpDados_Texto, tpNulo_NaoAceita)
  Set VerificaProdutoCardapioEst = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCardapio.VerificaProdutoCardapioEst]", _
            Err.Description
End Function



