VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLeituraMaquinaFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirLeituraMaquinaFecha(lngLEITURAMAQUINAFECHAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM LEITURAMAQUINAFECHA " & _
    "WHERE LEITURAMAQUINAFECHA.PKID = " & _
    Formata_Dados(lngLEITURAMAQUINAFECHAID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLeituraMaquinaFecha.ExcluirLeituraMaquinaFecha]", _
            Err.Description
End Function


Public Function SelecionarLeituraMaquinaFechaPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT LEITURAMAQUINAFECHA.* " & _
    "FROM LEITURAMAQUINAFECHA " & _
    "WHERE LEITURAMAQUINAFECHA.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  Set SelecionarLeituraMaquinaFechaPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLeituraMaquinaFecha.SelecionarLeituraMaquinaFechaPeloPkid]", _
            Err.Description
End Function

Public Sub InserirLeituraMaquinaFecha(lngPKID As Long, _
                                      lngLEITURAFECHAID As Long, _
                                      lngMAQUINAID As Long, _
                                      strMedicaoEntrada As String, _
                                      strMedicaoSaida As String, _
                                      strUsuario As String)

  On Error GoTo trata
  Dim strSql  As String
  Dim objRs   As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO LEITURAMAQUINAFECHA (LEITURAFECHAID, MAQUINAID, MEDICAOENTRADA, MEDICAOSAIDA, USUARIO) VALUES (" & _
    Formata_Dados(lngLEITURAFECHAID, tpDados_Longo) & _
    ", " & Formata_Dados(lngMAQUINAID, tpDados_Longo) & _
    ", " & Formata_Dados(strMedicaoEntrada, tpDados_Moeda) & _
    ", " & Formata_Dados(strMedicaoSaida, tpDados_Moeda) & _
    ", " & Formata_Dados(strUsuario, tpDados_Texto) & _
    ");"
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngPKID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLeituraMaquinaFecha.InserirLeituraMaquinaFecha]", _
            Err.Description
End Sub





