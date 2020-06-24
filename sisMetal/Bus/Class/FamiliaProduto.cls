VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFamiliaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirFamiliaProduto(strDescricao As String, _
                                 lngIPIID As Long, _
                                 lngICMSID As Long)

  On Error GoTo trata
  Dim objFamiliaProduto  As datSisMetal.clsFamiliaProduto
  '
  Set objFamiliaProduto = New datSisMetal.clsFamiliaProduto
  
  'Inserir Família de produtos
  objFamiliaProduto.InserirFamiliaProduto strDescricao, _
                                          lngIPIID, _
                                          lngICMSID
  '
  Set objFamiliaProduto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarFamiliaProduto(lngFAMILIAPRODUTOSID As Long, _
                                 strDescricao As String, _
                                 lngIPIID As Long, _
                                 lngICMSID As Long)
                         
  On Error GoTo trata
  Dim objFamiliaProduto As datSisMetal.clsFamiliaProduto
  Set objFamiliaProduto = New datSisMetal.clsFamiliaProduto
  
  'Alterar Família de produtos
  objFamiliaProduto.AlterarFamiliaProduto lngFAMILIAPRODUTOSID, _
                                          strDescricao, _
                                          lngIPIID, _
                                          lngICMSID

  '
  Set objFamiliaProduto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarFamiliaProduto(lngFAMILIAPRODUTOSID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objFamiliaProduto As datSisMetal.clsFamiliaProduto
  Set objFamiliaProduto = New datSisMetal.clsFamiliaProduto
  
  'Listar Família de produtos
  Set ListarFamiliaProduto = objFamiliaProduto.ListarFamiliaProduto(lngFAMILIAPRODUTOSID)
  '
  Set objFamiliaProduto = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirFamiliaProduto(ByVal lngFAMILIAPRODUTOSID As Long)
  Dim objFamiliaProduto    As datSisMetal.clsFamiliaProduto
  '
  On Error GoTo trata
  Set objFamiliaProduto = New datSisMetal.clsFamiliaProduto
  '
  objFamiliaProduto.ExcluirFamiliaProduto lngFAMILIAPRODUTOSID
  '
  Set objFamiliaProduto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


