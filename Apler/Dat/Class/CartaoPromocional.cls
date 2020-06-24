VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCartaoPromocional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub AlterarCartaoPromocional(lngCARTAOPROMOCIONALID As Long, _
                                    lngNumero As Long, _
                                    lngFXCARTAOPROMOCIONALID As Long, _
                                    DtRecebimento As String)
  On Error GoTo trata
  '
  Dim strSql      As String
  
  strSql = "UPDATE CARTAOPROMOCIONAL SET " & _
    "NUMERO = " & Formata_Dados(lngNumero, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    "FXCARTAOPROMOCIONALID = " & Formata_Dados(lngFXCARTAOPROMOCIONALID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    "DTRECEBIMENTO = " & Formata_Dados(DtRecebimento, tpDados_DataHora, tpNulo_Aceita) & _
    " WHERE PKID = " & Formata_Dados(lngCARTAOPROMOCIONALID, tpDados_Longo, tpNulo_NaoAceita)
    
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoPromocional.AlterarCartaoPromocional]", _
            Err.Description
End Sub

Public Sub InserirCartaoPromocional(lngNumero As Long, _
                                    lngFXCARTAOPROMOCIONALID As Long, _
                                    DtRecebimento As String)
  On Error GoTo trata
  '
  Dim strSql      As String
  
  strSql = "INSERT INTO CARTAOPROMOCIONAL (NUMERO, FXCARTAOPROMOCIONALID, DTRECEBIMENTO) VALUES (" & _
    Formata_Dados(lngNumero, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngFXCARTAOPROMOCIONALID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(DtRecebimento, tpDados_DataHora, tpNulo_Aceita) & ");"
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoPromocional.InserirCartaoPromocional]", _
            Err.Description
End Sub

Public Function ListarFxCartaoPromocionalPelaFaixa(lngNumero As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT FXCARTAOPROMOCIONAL.* FROM FXCARTAOPROMOCIONAL WHERE " & _
    Formata_Dados(lngNumero, tpDados_Longo, tpNulo_NaoAceita) & _
    " >= FXINICIAL AND " & Formata_Dados(lngNumero, tpDados_Longo, tpNulo_NaoAceita) & _
    " <= FXFINAL"
    
  Set ListarFxCartaoPromocionalPelaFaixa = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoPromocional.ListarFxCartaoPromocionalPelaFaixa]", _
            Err.Description
End Function

Public Function ListarCartaoPromocional(lngNumero As Long, _
                                        lngCARTAOPROMOCIONALID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CARTAOPROMOCIONAL.* FROM CARTAOPROMOCIONAL WHERE CARTAOPROMOCIONAL.NUMERO = " & _
    Formata_Dados(lngNumero, tpDados_Longo, tpNulo_NaoAceita) & _
    " AND CARTAOPROMOCIONAL.PKID <> " & Formata_Dados(lngCARTAOPROMOCIONALID, tpDados_Longo, tpNulo_NaoAceita)
    
  Set ListarCartaoPromocional = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoPromocional.ListarCartaoPromocional]", _
            Err.Description
End Function

Public Function ListarCartaoPromocionalPeloID(lngCARTAOPROMOCIONALID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT CARTAOPROMOCIONAL.* FROM CARTAOPROMOCIONAL WHERE CARTAOPROMOCIONAL.PKID = " & _
    Formata_Dados(lngCARTAOPROMOCIONALID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarCartaoPromocionalPeloID = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoPromocional.ListarCartaoPromocionalPeloID]", _
            Err.Description
End Function

Public Sub AlterarFxCartaoPromocional(lngFXCARTAOPROMOCIONALID As Long, _
                                      FxInicial As Long, _
                                      FxFinal As Long, _
                                      PercDesc As String, _
                                      lngCLIENTECARTAOPROMOCIONALID As Long, _
                                      Validade As String)
  On Error GoTo trata
  '
  Dim strSql      As String
  
  strSql = "UPDATE FXCARTAOPROMOCIONAL SET " & _
    "FXINICIAL = " & Formata_Dados(FxInicial, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    "FXFINAL = " & Formata_Dados(FxFinal, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    "PERCDESCONTO = " & Formata_Dados(PercDesc, tpDados_Moeda, tpNulo_NaoAceita) & ", " & _
    "VALIDADE = " & Formata_Dados(Validade, tpDados_DataHora, tpNulo_Aceita) & ", " & _
    "CLIENTECARTAOPROMOCIONALID = " & Formata_Dados(lngCLIENTECARTAOPROMOCIONALID, tpDados_Longo, tpNulo_Aceita) & _
    " WHERE PKID = " & Formata_Dados(lngFXCARTAOPROMOCIONALID, tpDados_Longo, tpNulo_Aceita) & " "
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoPromocional.AlterarFxCartaoPromocional]", _
            Err.Description
End Sub

Public Function ListarClienteCartaoPromocionalPeloNome(strCliente As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CLIENTECARTAOPROMOCIONAL.* FROM CLIENTECARTAOPROMOCIONAL WHERE CLIENTECARTAOPROMOCIONAL.NOME = " & _
    Formata_Dados(strCliente, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarClienteCartaoPromocionalPeloNome = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoPromocional.ListarClienteCartaoPromocionalPeloNome]", _
            Err.Description
End Function

Public Sub InserirClienteCartaoPromocional(strNome As String, _
                                           strNumero)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO CLIENTECARTAOPROMOCIONAL (NOME, NUMERO) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(strNumero, tpDados_Texto, tpNulo_NaoAceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoPromocional.InserirClienteCartaoPromocional]", _
            Err.Description
End Sub

Public Sub AlterarClienteCartaoPromocional(lngCLIENTECARTAOPROMOCIONALID As Long, _
                                           strNome As String, _
                                           strNumero As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE CLIENTECARTAOPROMOCIONAL SET NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_NaoAceita) & ", " & _
    "NUMERO = " & Formata_Dados(strNumero, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngCLIENTECARTAOPROMOCIONALID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoPromocional.AlterarClienteCartaoPromocional]", _
            Err.Description
End Sub

Public Function ListarClienteCartaoPromocional(lngCLIENTECARTAOPROMOCIONALID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CLIENTECARTAOPROMOCIONAL.* FROM CLIENTECARTAOPROMOCIONAL WHERE CLIENTECARTAOPROMOCIONAL.PKID = " & _
    Formata_Dados(lngCLIENTECARTAOPROMOCIONALID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarClienteCartaoPromocional = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoPromocional.ListarClienteCartaoPromocional]", _
            Err.Description
End Function

Public Sub ExlcuirClienteCartaoPromocional(lngCLIENTECARTAOPROMOCIONALID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM CLIENTECARTAOPROMOCIONAL WHERE PKID = " & _
    Formata_Dados(lngCLIENTECARTAOPROMOCIONALID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoPromocional.ExlcuirClienteCartaoPromocional]", _
            Err.Description
End Sub


Public Function ListarFxCartaoPromocional(lngFXCARTAOPROMOCIONALID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT FXCARTAOPROMOCIONAL.*, CLIENTECARTAOPROMOCIONAL.NOME FROM CLIENTECARTAOPROMOCIONAL LEFT JOIN FXCARTAOPROMOCIONAL ON CLIENTECARTAOPROMOCIONAL.PKID = FXCARTAOPROMOCIONAL.CLIENTECARTAOPROMOCIONALID WHERE FXCARTAOPROMOCIONAL.PKID = " & _
    Formata_Dados(lngFXCARTAOPROMOCIONALID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarFxCartaoPromocional = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoPromocional.ListarFxCartaoPromocional]", _
            Err.Description
End Function

Public Function ValidaFxCartaoPromocional(NumeroInicial As Long, _
                                          NumeroFinal As Long, _
                                          lngFXCARTAOPROMOCIONALID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT * FROM FXCARTAOPROMOCIONAL WHERE (FXINICIAL <= " & NumeroFinal & _
    " and FXFINAL >= " & NumeroInicial & ") AND PKID <> " & lngFXCARTAOPROMOCIONALID
  
  Set ValidaFxCartaoPromocional = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoPromocional.ValidaFxCartaoPromocional]", _
            Err.Description
End Function


Public Sub InserirFxCartaoPromocional(FxInicial As Long, _
                                      FxFinal As Long, _
                                      PercDesc As String, _
                                      lngCLIENTECARTAOPROMOCIONALID As Long, _
                                      Validade As String)
  On Error GoTo trata
  '
  Dim strSql      As String
  
  strSql = "INSERT INTO FXCARTAOPROMOCIONAL (FXINICIAL, FXFINAL, PERCDESCONTO, VALIDADE, CLIENTECARTAOPROMOCIONALID) VALUES (" & _
    Formata_Dados(FxInicial, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(FxFinal, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(PercDesc, tpDados_Moeda, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(Validade, tpDados_DataHora, tpNulo_Aceita) & ", " & _
    Formata_Dados(lngCLIENTECARTAOPROMOCIONALID, tpDados_Longo, tpNulo_Aceita) & ");"
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoPromocional.InserirFxCartaoPromocional]", _
            Err.Description
End Sub

Public Sub ExlcuirCartaoPromocional(lngCARTAOPROMOCIONALID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM CARTAOPROMOCIONAL WHERE PKID = " & _
    Formata_Dados(lngCARTAOPROMOCIONALID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoPromocional.ExlcuirCartaoPromocional]", _
            Err.Description
End Sub

Public Sub ExlcuirFxCartaoPromocional(lngFAIXACARTAOPROMOCIONALID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM FXCARTAOPROMOCIONAL WHERE PKID = " & _
    Formata_Dados(lngFAIXACARTAOPROMOCIONALID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCartaoPromocional.ExlcuirFxCartaoPromocional]", _
            Err.Description
End Sub


