VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Sub ExcluirCLIENTE(lngCLIENTEID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui CHEQUE
  strSql = "DELETE FROM CLIENTE WHERE CLIENTE.PKID = " & _
    Formata_Dados(lngCLIENTEID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCheque.ExcluirCLIENTE]", _
            Err.Description
End Sub

Public Sub ExcluirCHEQUE(lngCHEQUEID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  'exclui CHEQUE
  strSql = "DELETE FROM CHEQUE WHERE CHEQUE.PKID = " & _
    Formata_Dados(lngCHEQUEID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCheque.ExcluirCHEQUE]", _
            Err.Description
End Sub


Public Sub InserirCheque(lngCLIENTEID As Long, _
                         lngBANCOID As Long, _
                         lngMOTIVODEVOLID As Long, _
                         strConta As String, _
                         strCheque As String, _
                         strAgencia As String, _
                         strValor As String, _
                         strDtDevolucao As String, _
                         strDtRecebimento As String, _
                         strDtRecuperacao As String, _
                         strStatus As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO CHEQUE (CLIENTEID, BANCOID, CONTA, CHEQUE, AGENCIA, VALOR, " & _
    "DTDEVOLUCAO, MOTIVODEVOLID, DTRECEBIMENTO, DTRECUPERACAO, STATUS) values (" & _
    Formata_Dados(lngCLIENTEID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngBANCOID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strConta, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strCheque, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strAgencia, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strValor, tpDados_Moeda, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strDtDevolucao, tpDados_DataHora, tpNulo_Aceita) & ", " & _
    Formata_Dados(IIf(lngMOTIVODEVOLID = 0, "", lngMOTIVODEVOLID), tpDados_Longo, tpNulo_Aceita) & ", " & _
    Formata_Dados(strDtRecebimento, tpDados_DataHora, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strDtRecuperacao, tpDados_DataHora, tpNulo_Aceita) & ", " & _
    Formata_Dados(strStatus, tpDados_Texto, tpNulo_NaoAceita) & ");"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCheque.InserirCheque]", _
            Err.Description
End Sub

Public Sub AlterarCheque(lngCLIENTEID As Long, _
                         lngCHEQUEID As Long, _
                         lngBANCOID As Long, _
                         lngMOTIVODEVOLID As Long, _
                         strConta As String, _
                         strCheque As String, _
                         strAgencia As String, _
                         strValor As String, _
                         strDtDevolucao As String, _
                         strDtRecebimento As String, _
                         strDtRecuperacao As String, _
                         strStatus As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE CHEQUE SET CLIENTEID = " & Formata_Dados(lngCLIENTEID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    "BANCOID = " & Formata_Dados(lngBANCOID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    "CONTA = " & Formata_Dados(strConta, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "CHEQUE = " & Formata_Dados(strCheque, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "AGENCIA = " & Formata_Dados(strAgencia, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "VALOR = " & Formata_Dados(strValor, tpDados_Moeda, tpNulo_NaoAceita) & ", " & _
    "DTDEVOLUCAO = " & Formata_Dados(strDtDevolucao, tpDados_DataHora, tpNulo_Aceita) & ", " & _
    "MOTIVODEVOLID = " & Formata_Dados(IIf(lngMOTIVODEVOLID = 0, "", lngMOTIVODEVOLID), tpDados_Longo, tpNulo_Aceita) & ", " & _
    "DTRECEBIMENTO = " & Formata_Dados(strDtRecebimento, tpDados_DataHora, tpNulo_NaoAceita) & ", " & _
    "DTRECUPERACAO = " & Formata_Dados(strDtRecuperacao, tpDados_DataHora, tpNulo_Aceita) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngCHEQUEID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCheque.AlterarCheque]", _
            Err.Description
End Sub

Public Function ListarChequePorCPF(strCPF As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CHEQUE.* FROM CLIENTE INNER JOIN CHEQUE ON CLIENTE.PKID = CHEQUE.CLIENTEID " & _
    "WHERE CLIENTE.CPF = " & _
    Formata_Dados(strCPF, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarChequePorCPF = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCheque.ListarChequePorCPF]", _
            Err.Description
End Function

Public Function ListarCheque(lngCHEQUEID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CHEQUE.*, BANCO.NUMERO, BANCO.NOME, MOTIVODEVOL.CODMOTIVO, MOTIVODEVOL.DESCMOTIVO FROM (CHEQUE LEFT JOIN BANCO ON BANCO.PKID = CHEQUE.BANCOID) " & _
    "LEFT JOIN MOTIVODEVOL ON MOTIVODEVOL.PKID = CHEQUE.MOTIVODEVOLID " & _
    "WHERE CHEQUE.PKID = " & _
    Formata_Dados(lngCHEQUEID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarCheque = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCheque.ListarCheque]", _
            Err.Description
End Function

Public Function ListarClientePorCPF(strCPF As String, _
                                    lngCLIENTEID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT * FROM CLIENTE WHERE CPF = " & _
    Formata_Dados(strCPF, tpDados_Texto, tpNulo_NaoAceita) & _
    " AND PKID <> " & lngCLIENTEID
  Set ListarClientePorCPF = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCliente.ListarClientePorCPF]", _
            Err.Description
End Function

Public Sub InserirCliente(strCPF As String, _
                          strPlaca As String, _
                          strVeiculo As String, _
                          strNome As String, _
                          strTel1 As String, _
                          strTel2 As String, _
                          strTel3 As String, _
                          strEndereco As String, _
                          strCidade As String, _
                          strBairro As String, _
                          strDtNasc As String, _
                          strEstado As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO CLIENTE (CPF, PLACA, VEICULO, NOME, TEL1, TEL2, TEL3, ENDERECO, " & _
    "CIDADE, BAIRRO, DTNASCIMENTO, ESTADO) VALUES (" & _
    Formata_Dados(strCPF, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strPlaca, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strVeiculo, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strTel1, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strTel2, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strTel3, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strEndereco, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strCidade, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strBairro, tpDados_Texto, tpNulo_Aceita) & ", " & _
    Formata_Dados(strDtNasc, tpDados_DataHora, tpNulo_Aceita) & ", " & _
    Formata_Dados(strEstado, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCheque.InserirCliente]", _
            Err.Description
End Sub

Public Sub AlterarCliente(lngCLIENTEID As Long, _
                          strCPF As String, _
                          strPlaca As String, _
                          strVeiculo As String, _
                          strNome As String, _
                          strTel1 As String, _
                          strTel2 As String, _
                          strTel3 As String, _
                          strEndereco As String, _
                          strCidade As String, _
                          strBairro As String, _
                          strDtNasc As String, _
                          strEstado As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE CLIENTE SET CPF = " & Formata_Dados(strCPF, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "PLACA = " & Formata_Dados(strPlaca, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "VEICULO = " & Formata_Dados(strVeiculo, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "TEL1 = " & Formata_Dados(strTel1, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "TEL2 = " & Formata_Dados(strTel2, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "TEL3 = " & Formata_Dados(strTel3, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "ENDERECO = " & Formata_Dados(strEndereco, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "CIDADE = " & Formata_Dados(strCidade, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "BAIRRO = " & Formata_Dados(strBairro, tpDados_Texto, tpNulo_Aceita) & ", " & _
    "DTNASCIMENTO = " & Formata_Dados(strDtNasc, tpDados_DataHora, tpNulo_Aceita) & ", " & _
    "ESTADO = " & Formata_Dados(strEstado, tpDados_Texto, tpNulo_Aceita) & " " & _
    "WHERE PKID = " & lngCLIENTEID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCheque.AlterarCliente]", _
            Err.Description
End Sub

Public Function ListarCliente(lngCLIENTEID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT * FROM CLIENTE WHERE PKID = " & _
    Formata_Dados(lngCLIENTEID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarCliente = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCliente.ListarCliente]", _
            Err.Description
End Function



