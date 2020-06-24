VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEntradaMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirITEMENTRADA(ByVal lngITEMENTRADAID As Long, _
                              ByVal lngESTOQUEID As Long, _
                              ByVal lngQUANTIDADENOVA As Long)
  Dim clsEntMat    As datSisLoc.clsEntradaMaterial
  '
  On Error GoTo trata
  Set clsEntMat = New datSisLoc.clsEntradaMaterial
  '
  clsEntMat.ExcluirITEMENTRADA lngITEMENTRADAID, _
                                       lngESTOQUEID, _
                                       lngQUANTIDADENOVA
  '
  Set clsEntMat = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirEntradaMaterial(strDescricao As String, _
                                  strUsuario As String, _
                                  lngDOCUMENTOID As Long, _
                                  strFornecedor As String, _
                                  strDataAquisicao As String, _
                                  strNumero As String, _
                                  ByRef lngENTRADAMATERIALID As Long, _
                                  ByRef strData As String, _
                                  ByRef strCodigo As String)
  On Error GoTo trata
  
  Dim clsEntMat   As datSisLoc.clsEntradaMaterial
  Dim strSql      As String
  Dim objRs       As ADODB.Recordset
  '
  strData = Format(Now, "DD/MM/YYYY hh:mm")
  '
  Set clsEntMat = New datSisLoc.clsEntradaMaterial
  strCodigo = clsEntMat.CapturaCodigoEntradaMaterial
  'Inserir Entrada de material
  clsEntMat.InserirEntradaMaterial strData, _
                                   strDescricao, _
                                   strCodigo, _
                                   strUsuario, _
                                   lngDOCUMENTOID, _
                                   strFornecedor, _
                                   strDataAquisicao, _
                                   strNumero, _
                                   lngENTRADAMATERIALID
  '
  'Após inserir a entrada de material, Retornar PKID pelo código
  '
  Set clsEntMat = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Function ValidarExclusaoEntradaMaterial(lngENTRADAMATERIALID As Long) As String
  On Error GoTo trata
  Dim strSql      As String
  Dim strRetorno  As String
  Dim objRs       As ADODB.Recordset
  Dim clsGer      As datSisLoc.clsGeral
  '
  Set clsGer = New datSisLoc.clsGeral
  strRetorno = ""
  'Valida entrada de material
  'ITEMENTRADA
  strSql = "Select count(*) AS TOTAL from ITEMENTRADA WHERE ENTRADAMATERIALID = " & lngENTRADAMATERIALID
  Set objRs = clsGer.ExecutarSQL(strSql)
  '
  If Not objRs.EOF Then
    If objRs.Fields("TOTAL").Value > 0 Then
      strRetorno = "Itens de entrada de material"
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  '
  ValidarExclusaoEntradaMaterial = strRetorno
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub AlterarEntradaMaterial(lngENTRADAMATERIALID As Long, _
                                  strDescricao As String, _
                                  lngDOCUMENTOID As Long, _
                                  strFornecedor As String, _
                                  strNumero As String, _
                                  strDataAquisicao As String)
                         
  On Error GoTo trata
  Dim clsEntMat As datSisLoc.clsEntradaMaterial
  Set clsEntMat = New datSisLoc.clsEntradaMaterial
  
  'Inserir Entrada de Material
  clsEntMat.AlterarEntradaMaterial lngENTRADAMATERIALID, _
                                   strDescricao, _
                                   lngDOCUMENTOID, _
                                   strFornecedor, _
                                   strNumero, _
                                   strDataAquisicao

  '
  Set clsEntMat = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarEstoquePelaEntradaMaterial(ByVal lngESTOQUEID As Long, _
                                             ByVal strCodigo As String, _
                                             ByVal strQuantidade As String)
  '
  On Error GoTo trata
    
  Dim clsEntMat As datSisLoc.clsEntradaMaterial
  '
  Set clsEntMat = New datSisLoc.clsEntradaMaterial
  '
  clsEntMat.AlterarEstoquePelaEntradaMaterial lngESTOQUEID, _
                                              strCodigo, _
                                              strQuantidade
  '
  Set clsEntMat = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirITEMENTRADA(ByVal lngENTRADAMATERIALID As Long, _
                                      ByVal lngESTOQUEID As Long, _
                                      ByVal strQuantidade As String)
  '
  On Error GoTo trata
    
  Dim clsEntMat As datSisLoc.clsEntradaMaterial
  '
  Set clsEntMat = New datSisLoc.clsEntradaMaterial
  '
  clsEntMat.InserirITEMENTRADA lngENTRADAMATERIALID, _
                                       lngESTOQUEID, _
                                       strQuantidade
  '
  Set clsEntMat = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarEntradaMaterial(lngENTRADAMATERIALID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsEntMat As datSisLoc.clsEntradaMaterial
  Set clsEntMat = New datSisLoc.clsEntradaMaterial
  
  'Listar entrada de material
  Set ListarEntradaMaterial = clsEntMat.ListarEntradaMaterial(lngENTRADAMATERIALID)
  '
  Set clsEntMat = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub ExcluirENTRADAMATERIAL(ByVal lngENTRADAMATERIALID As Long)
  Dim clsEntMat    As datSisLoc.clsEntradaMaterial
  '
  On Error GoTo trata
  Set clsEntMat = New datSisLoc.clsEntradaMaterial
  '
  clsEntMat.ExcluirENTRADAMATERIAL lngENTRADAMATERIALID
  '
  Set clsEntMat = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


