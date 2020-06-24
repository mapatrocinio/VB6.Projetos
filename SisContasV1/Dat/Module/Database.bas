Attribute VB_Name = "Database"
Option Explicit
Option Compare Text

' Sleep API is declared in the form to keep the
' SetWaitableTimer code in its own re-usable module.
Private Declare Sub Sleep Lib "kernel32" ( _
    ByVal dwMilliseconds As Long)

Public Const MAX_RETRIES = 10 '50 seg
Public Const psSenhaPerm = "SHOGUM2806"
Public Const strClassName As String = "DatSisContas"
Public Const lngCteErroData As Long = vbObjectError + 513
Const lngTimeout As Long = 120

Enum TpTipoDados
  tpDados_Texto '0 a 255
  tpDados_Memo 'Sem Limite
  tpDados_Inteiro '-32767 a 32767
  tpDados_Longo 'Sem limite
  tpDados_DataHora 'MM/DD/YYYY hh:mm:ss
  tpDados_Moeda '121212.98
  tpDados_Boolean '121212.98
End Enum

Enum tpAceitaNulo
  tpNulo_Aceita
  tpNulo_NaoAceita
End Enum
Public Sub TratarErro(ByVal pNumero As Long, _
                    ByVal pDescricao As String, _
                    ByVal pSource As String)
  '
  On Error Resume Next
  Dim strUsuario As String
  Dim intI       As Integer

  intI = FreeFile
  Open App.Path & "\SisConta.txt" For Append As #intI

  Print #intI, Format(Now(), "DD/MM/YYYY hh:mm") & ";" & pNumero & ";" & pDescricao & ";" & pSource
  Close #intI
  'mostrar Mensagem
  'MsgBox "O Seginte Erro Ocorreu: " & vbCrLf & vbCrLf & _
    "Número: " & pNumero & vbCrLf & _
    "Descrição: " & pDescricao & vbCrLf & vbCrLf & _
    "Módulo: " & pSource & vbCrLf & _
    "Data/Hora: " & Format(Now(), "DD/MM/YYYY hh:mm") & vbCrLf & _
    "Erro gravado no arquivo: " & App.Path & "\SisMotel.txt" & vbCrLf & vbCrLf & _
    "Caso o erro persista contacte o suporte e envie o arquivo acima, informando a data e hora acima informada da ocorrência deste erro.", vbCritical, TITULOSISTEMA
End Sub

Public Function RetonaProximoSequencial(strTabela As String, _
                                        strCampo As String, _
                                        Optional strWhere As String) As Variant
  On Error GoTo trata
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim lngRet    As Long
  strSql = "SELECT MAX(" & strCampo & ") as maximo FROM " & strTabela
  If strWhere & "" <> "" Then
    strSql = strSql & " WHERE " & strWhere
  End If
  Set objRs = RunSPReturnRS(strSql)
  If objRs.EOF Then
    lngRet = 1
  ElseIf Not IsNumeric(objRs.Fields("maximo").Value) Then
    lngRet = 1
  Else
    lngRet = objRs.Fields("maximo").Value + 1
  End If
  RetonaProximoSequencial = lngRet
  Exit Function
trata:
  Err.Raise Err.Number, "[Database.RetonaProximoSequencial]", Err.Description
End Function

Public Function Tira_Plic(pValor As String) As Variant
  Tira_Plic = Replace(pValor, "'", "''")
End Function

Function GetConnectionString() As String
On Error GoTo errorHandler
  'GetConnectionString = "Provider=SQLOLEDB" & _
    ";Initial Catalog=ViaGlobal_Novo" & _
    ";Data Source=Web" & _
    ";User Id=SA" & _
    ";Password="
  'GetConnectionString = "DSN=SisMotel" & _
    ";User Id=Admin" & _
    ";Password=SHOGUM2806"
  GetConnectionString = "DSN=SisContas" & _
    ";Password=" & psSenhaPerm


    
  Exit Function
errorHandler:
    TrataErro strClassName, "GetConnectionString", Err.Number, Err.Description
End Function

Function RunSPReturnRS(ByVal strSP As String, _
                        ParamArray params() As Variant) As ADODB.Recordset

  On Error GoTo errorHandler
  
  ' Create the ADO objects
  Dim rs  As ADODB.Recordset
  Dim cmd As ADODB.Command
  Dim lngErrCount     As Long

  lngErrCount = 0
  
  Set rs = CriaInstancia("ADODB.Recordset")
  Set cmd = CriaInstancia("ADODB.Command")
  
  ' Init the ADO objects  & the stored proc parameters
  cmd.ActiveConnection = GetConnectionString
  cmd.CommandTimeout = lngTimeout
  cmd.CommandText = strSP
  cmd.CommandType = adCmdText
  
  collectParams cmd, params
  
  ' Execute the query for readonly
  rs.CursorLocation = adUseClient
  rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
  
  ' Disconnect the recordset
  Set cmd.ActiveConnection = Nothing
  Set cmd = Nothing
  Set rs.ActiveConnection = Nothing
  
  ' Return the resultant recordset
  Set RunSPReturnRS = rs

  Exit Function
errorHandler:
  If Err.Number = -2147467259 Then
    'ATENÇAO, OCORREU O ERRO
    lngErrCount = lngErrCount + 1
    If lngErrCount < MAX_RETRIES Then
      Sleep 5000 'sleep 5 seconds
      TratarErro Err.Number, _
                 lngErrCount & " - Erro tratado na Função Sleep", _
                 "[" & strClassName & "Database.RunSPReturnRS]"
      Resume
    Else
      'Retries did not help. trate error
      Err.Raise Err.Number, _
                "[" & strClassName & "Database.RunSPReturnRS]", _
                Err.Description & " | sql: " & strSP
    End If
  Else
    Err.Raise Err.Number, _
              "[" & strClassName & "Database.RunSPReturnRS]" & _
              Err.Description & " | sql: " & strSP
  End If
End Function

Function RunSP(ByVal strSP As String, ParamArray params() As Variant)
  On Error GoTo errorHandler

  ' Create the ADO objects
  Dim cmd             As ADODB.Command
  Dim lngErrCount     As Long
  lngErrCount = 0
  Set cmd = CriaInstancia("ADODB.Command")
  

  ' Init the ADO objects & the stored proc parameters
  cmd.ActiveConnection = GetConnectionString
  cmd.CommandTimeout = lngTimeout
  cmd.CommandText = strSP
  cmd.CommandType = adCmdText
  collectParams cmd, params

  ' Execute the query without returning a recordset
  cmd.Execute , , adExecuteNoRecords
  ' Disconnect the recordset and clean up
  Set cmd.ActiveConnection = Nothing
  Set cmd = Nothing

  Exit Function

errorHandler:
  If Err.Number = -2147467259 Then
    'ATENÇAO, OCORREU O ERRO
    lngErrCount = lngErrCount + 1
    If lngErrCount < MAX_RETRIES Then
      Sleep 5000 'sleep 5 seconds
      TratarErro Err.Number, _
                 lngErrCount & " - Erro tratado na Função Sleep", _
                 "[" & strClassName & "Database.RunSP]"
      Resume
    Else
      'Retries did not help. trate error
      Err.Raise Err.Number, _
                "[" & strClassName & "Database.RunSP]", _
                Err.Description & " | sql: " & strSP
    End If
  Else
    Err.Raise Err.Number, _
              "[" & strClassName & "Database.RunSP]", _
              Err.Description & " | sql: " & strSP
  End If
End Function

''Function RunSPReturnInteger(ByVal strSP As String, ParamArray params() As Variant) As Long 'adInterger is really a VB Long
''    On Error GoTo errorHandler
''
''    ' Create the ADO objects
''    Dim cmd As ADODB.Command
''    Set cmd = CriaInstancia("ADODB.Command")
''
''    ' Init the ADO objects & the stored proc parameters
''    cmd.ActiveConnection = GetConnectionString
''    cmd.CommandTimeout = lngTimeout
''    cmd.CommandText = strSP
''    cmd.CommandType = adCmdStoredProc
''    cmd.Parameters.Append cmd.CreateParameter("@retval", adInteger, adParamReturnValue, 4)
''    collectParams cmd, params
''
''    ' Assume the last parameter is outgoing
''
''
''    ' Execute without a resulting recordset and pull out the "return value" parameter
''    cmd.Execute , , adExecuteNoRecords
''    RunSPReturnInteger = cmd.Parameters("@retval").Value
''
''    ' Disconnect the recordset, and clean up
''    Set cmd.ActiveConnection = Nothing
''    Set cmd = Nothing
''
''    Exit Function
''
''errorHandler:
''  TrataErro strClassName, "RunSPReturnInteger", Err.Number, Err.Description
''End Function

Private Sub collectParams(ByRef cmd As ADODB.Command, _
                          ParamArray argparams() As Variant)
                          
  Dim params As Variant, v As Variant
  Dim i As Integer, L As Integer, u As Integer
  Dim oParam As ADODB.Parameter
On Error GoTo errorHandler
  params = argparams(0)
  For i = LBound(params) To UBound(params)
    L = LBound(params(i))
    u = UBound(params(i))
    ' Check for nulls.
    
    If VarType(params(i)(3)) = vbString Then
        v = IIf(params(i)(3) = "", Null, params(i)(3))
    Else
        v = params(i)(3)
    End If
    Set oParam = cmd.CreateParameter(params(i)(0), params(i)(1), adParamInput, params(i)(2), v)
    If oParam.Type = adNumeric Then
        oParam.Precision = params(i)(4)
        oParam.NumericScale = params(i)(5)
    End If
    cmd.Parameters.Append oParam
  Next i
  Exit Sub
errorHandler:
  TrataErro strClassName, "collectParams", Err.Number, Err.Description
End Sub

'Change the transaction state, and raise an error.
Public Sub CtxTrataErro(NomeModulo As String, NomeFuncao As String)
   
   Err.Raise Err.Number, GetErrSourceDescription(NomeModulo, NomeFuncao), Err.Description
    
End Sub



Private Function GetErrSourceDescription(modName As String, _
                                          procName As String) As String
  GetErrSourceDescription = "[" & modName & "." & procName & "]"
End Function



Public Sub TrataErro(strNomeModulo As String, _
                      strNomeFuncao As String, _
                      lngNumeroErro As Long, _
                      strDescErro As String)

  LogErro lngNumeroErro, strDescErro, strNomeModulo, strNomeFuncao
  'transporta para fora erro
  Err.Raise lngNumeroErro, GetErrSourceDescription(strNomeModulo, strNomeFuncao), strDescErro

End Sub

Private Function CriaInstancia(ByVal pClasseObjeto As String) As Object
  Set CriaInstancia = CreateObject(pClasseObjeto)
End Function


Public Sub LogErro(ByVal pNumero As Long, _
                    ByVal pDescricao As String, _
                    ByVal pModulo As String, _
                    ByVal pFuncao As String)
                    
    On Error Resume Next
    Dim strUsuario As String
    Dim intI       As Integer
        
    intI = FreeFile
    Open App.Path & "\SisContas.txt" For Append As #intI
    
    Print #intI, Format(Now(), "DD/MM/YYYY hh:mm") & ";" & pModulo & "." & pFuncao & ";" & pNumero & ";" & pDescricao
    Close #intI
End Sub

Public Function Formata_Dados(pValor As Variant, pTipoDados As TpTipoDados, pAceitaNulo As tpAceitaNulo, Optional pTamanhoCampo As Integer) As Variant
  On Error GoTo trata
  '
  Dim vRetorno As Variant
  Dim sData As String
  '
  Select Case pTipoDados
  Case TpTipoDados.tpDados_Boolean
    If pValor Then
      vRetorno = "True"
    Else
      vRetorno = "False"
    End If
  Case TpTipoDados.tpDados_Texto
    If Len(Trim(pValor & "")) = 0 Then
      If pAceitaNulo = tpNulo_Aceita Then
        vRetorno = "Null"
      Else
        vRetorno = "' '"
      End If
    Else
      vRetorno = "'" & Tira_Plic(Trim(pValor & "")) & "'"
    End If
  Case TpTipoDados.tpDados_Longo
    If Not IsNumeric(pValor) Then
      If pAceitaNulo = tpNulo_Aceita Then
        vRetorno = "Null"
      Else
        vRetorno = "0"
      End If
    Else
      vRetorno = CLng(pValor)
    End If
  Case TpTipoDados.tpDados_DataHora
    'Converter para Data
    sData = ""
    If Len(pValor & "") = 10 Then
      If Mid(pValor & "", 1, 2) <> "__" Then
        'Data no Formato DD/MM/YYYY
        sData = "#" & Mid(pValor, 4, 2) & "/" & Mid(pValor, 1, 2) & "/" & Mid(pValor, 7, 4) & "#"
      Else
        sData = "null"
      End If
    ElseIf Len(pValor & "") = 16 Then
      'Data no Formato DD/MM/YYYY hh:mm
      If Mid(pValor & "", 1, 2) <> "__" Then
        'Data no Formato DD/MM/YYYY
        sData = "#" & Mid(pValor, 4, 2) & "/" & Mid(pValor, 1, 2) & "/" & Mid(pValor, 7, 4) & " " & Mid(pValor, 12, 2) & ":" & Mid(pValor, 15, 2) & "#"
      Else
        sData = "null"
      End If
    Else
      sData = ""
    End If
    If Len(sData) = 0 Then
      If pAceitaNulo = tpNulo_Aceita Then
        vRetorno = "Null"
      Else
        vRetorno = "#01/01/1900#"
      End If
    Else
      vRetorno = sData
    End If
  Case TpTipoDados.tpDados_Moeda
    If Not IsNumeric(pValor) Then
      If pAceitaNulo = tpNulo_Aceita Then
        vRetorno = "Null"
      Else
        vRetorno = "0"
      End If
    Else
      vRetorno = Replace(pValor, ".", "")
      vRetorno = Replace(vRetorno, ",", ".")
    End If
  Case Else
  End Select
  '
  Formata_Dados = vRetorno
  '
  Exit Function
trata:
  
End Function






