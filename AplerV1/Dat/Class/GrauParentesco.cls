VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrauParentesco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirGrauParentesco(strDescricao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO GRAUPARENTESCO (DESCRICAO) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGrauParentesco.InserirGrauParentesco]", _
            Err.Description
End Sub

Public Sub AlterarGrauParentesco(lngGRAUPARENTESCOID As Long, _
                                 strDescricao As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE GRAUPARENTESCO SET DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngGRAUPARENTESCOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGrauParentesco.AlterarGrauParentesco]", _
            Err.Description
End Sub

Public Function ListarGrauParentesco(lngGRAUPARENTESCOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT GRAUPARENTESCO.* " & _
    "FROM GRAUPARENTESCO " & _
    "WHERE GRAUPARENTESCO.PKID = " & _
    Formata_Dados(lngGRAUPARENTESCOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarGrauParentesco = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGrauParentesco.ListarGrauParentesco]", _
            Err.Description
End Function

Public Sub ExlcuirGrauParentesco(lngGRAUPARENTESCOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM GRAUPARENTESCO WHERE PKID = " & _
    Formata_Dados(lngGRAUPARENTESCOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGrauParentesco.ExlcuirGrauParentesco]", _
            Err.Description
End Sub


Public Function ListarGrauParentescoPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT GRAUPARENTESCO.* " & _
    "FROM GRAUPARENTESCO " & _
    "WHERE GRAUPARENTESCO.DESCRICAO = " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarGrauParentescoPelaDesc = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGrauParentesco.ListarGrauParentescoPelaDesc]", _
            Err.Description
End Function

