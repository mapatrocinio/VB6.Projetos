VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlanoConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirPlanoConvenio(lngPlanoConvenioId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM PLANOCONVENIO " & _
    "WHERE PLANOCONVENIO.PKID = " & _
    Formata_Dados(lngPlanoConvenioId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPlanoConvenio.ExcluirPlanoConvenio]", _
            Err.Description
End Function

Public Function SelecionarPlanoConvenioPeloPkid(lngPlanoConvenioId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PLANOCONVENIO.* " & _
    "FROM PLANOCONVENIO " & _
    "WHERE PLANOCONVENIO.PKID = " & _
    Formata_Dados(lngPlanoConvenioId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarPlanoConvenioPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPlanoConvenio.SelecionarPlanoConvenioPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarPlanoConvenio(lngPKID As Long, _
                                strNome As String, _
                                strValorSocio As String, _
                                strValorDependente As String, _
                                strValorAplerSocio As String, _
                                strValorAplerDependente As String, _
                                strDataInicio As String, _
                                strDataFim As String, _
                                strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE PLANOCONVENIO SET " & _
    "NOME = " & Formata_Dados(strNome, tpDados_Texto) & ", " & _
    "VALORSOCIO = " & Formata_Dados(strValorSocio, tpDados_Moeda) & ", " & _
    "VALORDEPENDENTE = " & Formata_Dados(strValorDependente, tpDados_Moeda) & ", " & _
    "VALORAPLERSOCIO = " & Formata_Dados(strValorAplerSocio, tpDados_Moeda) & ", " & _
    "VALORAPLERDEPENDENTE = " & Formata_Dados(strValorAplerDependente, tpDados_Moeda) & ", " & _
    "DATAINICIO = " & Formata_Dados(strDataInicio, tpDados_DataHora) & ", " & _
    "DATAFIM = " & Formata_Dados(strDataFim, tpDados_DataHora) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPlanoConvenio.AlterarPlanoConvenio]", _
            Err.Description
End Sub

Public Sub InserirPlanoConvenio(lngCONVENIOID As Long, _
                                strNome As String, _
                                strValorSocio As String, _
                                strValorDependente As String, _
                                strValorAplerSocio As String, _
                                strValorAplerDependente As String, _
                                strDataInicio As String, _
                                strDataFim As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO PLANOCONVENIO (CONVENIOID, NOME, VALORSOCIO, VALORDEPENDENTE, VALORAPLERSOCIO, VALORAPLERDEPENDENTE, DATAINICIO, DATAFIM, STATUS) VALUES (" & _
    Formata_Dados(lngCONVENIOID, tpDados_Longo) & ", " & _
    Formata_Dados(strNome, tpDados_Texto) & ", " & _
    Formata_Dados(strValorSocio, tpDados_Moeda) & ", " & _
    Formata_Dados(strValorDependente, tpDados_Moeda) & ", " & _
    Formata_Dados(strValorAplerSocio, tpDados_Moeda) & ", " & _
    Formata_Dados(strValorAplerDependente, tpDados_Moeda) & ", " & _
    Formata_Dados(strDataInicio, tpDados_DataHora) & ", " & _
    Formata_Dados(strDataFim, tpDados_DataHora) & ", " & _
    Formata_Dados("A", tpDados_Texto) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPlanoConvenio.InserirPlanoConvenio]", _
            Err.Description
End Sub

Public Function SelecionarPlanoConvenio(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PLANOCONVENIO.* " & _
    "FROM PLANOCONVENIO " & _
    "WHERE PLANOCONVENIO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarPlanoConvenio = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPlanoConvenio.SelecionarPlanoConvenio]", _
            Err.Description
End Function

