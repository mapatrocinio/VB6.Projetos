VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReferenciaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function ListarReferenciaProdutoPelaDesc(strReferencia As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objReferenciaProduto As datSisMetal.clsReferenciaProduto
  Set objReferenciaProduto = New datSisMetal.clsReferenciaProduto
  
  'Listar Documento de entrada
  Set ListarReferenciaProdutoPelaDesc = objReferenciaProduto.ListarReferenciaProdutoPelaDesc(strReferencia)
  '
  Set objReferenciaProduto = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function ValidarExclusaoReferenciaProduto(lngREFERENCIAPRODUTOID As Long) As String
  On Error GoTo trata
  Dim strSql      As String
  Dim strRetorno  As String
  Dim objRs       As ADODB.Recordset
  Dim objGeral      As datSisMetal.clsGeral
  '
  Set objGeral = New datSisMetal.clsGeral
  strRetorno = ""
  'Valida entrada de material
  'TAB_REFERENCIAPRODUTO
  strSql = "Select count(*) AS TOTAL from GRUPO_PRODUTO WHERE REFERENCIAPRODUTOID = " & lngREFERENCIAPRODUTOID
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  If Not objRs.EOF Then
    If objRs.Fields("TOTAL").Value > 0 Then
      strRetorno = "Grupo de Produtos"
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  '
  ValidarExclusaoReferenciaProduto = strRetorno
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub InserirReferenciaProduto(strReferencia As String)

  On Error GoTo trata
  Dim objReferenciaProduto  As datSisMetal.clsReferenciaProduto
  '
  Set objReferenciaProduto = New datSisMetal.clsReferenciaProduto
  
  'Inserir Documento de entrada
  objReferenciaProduto.InserirReferenciaProduto strReferencia
  '
  Set objReferenciaProduto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarReferenciaProduto(lngREFERENCIAPRODUTOID As Long, _
                                    strReferencia As String)
                         
  On Error GoTo trata
  Dim objReferenciaProduto As datSisMetal.clsReferenciaProduto
  Set objReferenciaProduto = New datSisMetal.clsReferenciaProduto
  
  'Inserir Documento de Entrada
  objReferenciaProduto.AlterarReferenciaProduto lngREFERENCIAPRODUTOID, _
                                                strReferencia

  '
  Set objReferenciaProduto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Function ListarReferenciaProduto(lngREFERENCIAPRODUTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objReferenciaProduto As datSisMetal.clsReferenciaProduto
  Set objReferenciaProduto = New datSisMetal.clsReferenciaProduto
  
  'Listar Documento de entrada
  Set ListarReferenciaProduto = objReferenciaProduto.ListarReferenciaProduto(lngREFERENCIAPRODUTOID)
  '
  Set objReferenciaProduto = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirREFERENCIAPRODUTO(ByVal lngREFERENCIAPRODUTOID As Long)
  Dim objReferenciaProduto    As datSisMetal.clsReferenciaProduto
  '
  On Error GoTo trata
  Set objReferenciaProduto = New datSisMetal.clsReferenciaProduto
  '
  objReferenciaProduto.ExcluirREFERENCIAPRODUTO lngREFERENCIAPRODUTOID
  '
  Set objReferenciaProduto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

