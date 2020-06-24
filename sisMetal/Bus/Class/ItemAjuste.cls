VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItemAjuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirItemAjuste(lngITEM_AJUSTEID As Long, _
                             lngAJUSTEID As Long, _
                             lngPERFILID As Long, _
                             strQuantidade As String)
  On Error GoTo trata
  Dim objItemAjuste   As datSisMetal.clsItemAjuste
  '
  Set objItemAjuste = New datSisMetal.clsItemAjuste
  'Inserir ItemAjuste
  
  objItemAjuste.InserirItemAjuste lngITEM_AJUSTEID, _
                                  lngAJUSTEID, _
                                  lngPERFILID, _
                                  strQuantidade
  '
  Set objItemAjuste = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
'''
'''Public Function SelecionarItemAjustePeloPkid(lngItemAjusteId As Long) As ADODB.Recordset
'''  On Error GoTo trata
'''  Dim objItemAjuste As datSisMetal.clsItemAjuste
'''  Set objItemAjuste = New datSisMetal.clsItemAjuste
'''
'''  'Listar ItemAjuste
'''  Set SelecionarItemAjustePeloPkid = objItemAjuste.SelecionarItemAjustePeloPkid(lngItemAjusteId)
'''  '
'''  Set objItemAjuste = Nothing
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Function
'''
'''Public Sub ExcluirItemAjuste(lngItemAjusteId As Long)
'''  On Error GoTo trata
'''  Dim objItemAjuste As datSisMetal.clsItemAjuste
'''  Set objItemAjuste = New datSisMetal.clsItemAjuste
'''
'''  'Excluir ItemAjuste
'''  objItemAjuste.ExcluirItemAjuste lngItemAjusteId
'''  '
'''  Set objItemAjuste = Nothing
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Sub
'''
'''
'''Public Sub InserirItemAjuste(lngAJUSTEID As Long, _
'''                             lngLINHAID As Long, _
'''                             strPeso As String, _
'''                             strComprimentoVara As String)
'''  On Error GoTo trata
'''  Dim objItemAjuste       As datSisMetal.clsItemAjuste
'''  Dim curQuantidade       As Currency
'''  '
'''  curQuantidade = CalculoQuantidadeAjuste(lngLINHAID, _
'''                                          strPeso)
'''  '
'''  Set objItemAjuste = New datSisMetal.clsItemAjuste
'''  'Inserir ItemAjuste
'''  objItemAjuste.InserirItemAjuste lngAJUSTEID, _
'''                                  lngLINHAID, _
'''                                  Format(IIf(curQuantidade = 0, "", curQuantidade), "###,###") & "", _
'''                                  strPeso, _
'''                                  strComprimentoVara
'''  '
'''  Set objItemAjuste = Nothing
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Sub
'''
'''Public Function CalculoQuantidadeAjuste(lngLINHAID As Long, _
'''                                        strPeso As String) As Currency
'''
'''  On Error GoTo trata
'''  Dim objItemAjuste       As datSisMetal.clsItemAjuste
'''  'Obter peso da linha
'''  Set objItemAjuste = New datSisMetal.clsItemAjuste
'''  CalculoQuantidadeAjuste = objItemAjuste.CalculoQuantidadeAjuste(lngLINHAID, _
'''                                                                  strPeso)
'''  Set objItemAjuste = Nothing
'''  '
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Function
'''
'''
'''Public Sub AlterarItemAjuste(lngPKID As Long, _
'''                             lngLINHAID As Long, _
'''                             strPeso As String, _
'''                             strComprimentoVara As String)
'''
'''  On Error GoTo trata
'''  Dim objItemAjuste       As datSisMetal.clsItemAjuste
'''  Dim curQuantidade       As Currency
'''  '
'''  curQuantidade = CalculoQuantidadeAjuste(lngLINHAID, _
'''                                          strPeso)
'''  Set objItemAjuste = New datSisMetal.clsItemAjuste
'''  'Atualizar ItemAjuste
'''  objItemAjuste.AlterarItemAjuste lngPKID, _
'''                                  lngLINHAID, _
'''                                  Format(IIf(curQuantidade = 0, "", curQuantidade), "###,###") & "", _
'''                                  strPeso, _
'''                                  strComprimentoVara
'''  '
'''  Set objItemAjuste = Nothing
'''  Exit Sub
'''trata:
'''  Err.Raise Err.Number, Err.Source, Err.Description
'''End Sub
'''
