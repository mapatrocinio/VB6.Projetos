VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServicoFim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public Function SelecionarServicoFimPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT SERVICO.* " & _
    " FROM SERVICO " & _
    "WHERE SERVICO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarServicoFimPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsServicoFim.SelecionarServicoFimPeloPkid]", _
            Err.Description
End Function

Public Sub AlterarServicoFim(lngPKID As Long, _
                          strDataHora As String, _
                          strObservacao As String, _
                          strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE SERVICO SET " & _
    "DATAFINAL = " & Formata_Dados(strDataHora, tpDados_DataHora) & ", " & _
    "OBSERVACAOFINAL = " & Formata_Dados(strObservacao, tpDados_Texto) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsServicoFim.AlterarServicoFim]", _
            Err.Description
End Sub
