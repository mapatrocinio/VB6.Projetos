VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItemOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarItemOSPeloPkid(lngITEMOSID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objItemOS As datSisMetal.clsItemOS
  Set objItemOS = New datSisMetal.clsItemOS
  
  'Listar ItemOS
  Set SelecionarItemOSPeloPkid = objItemOS.SelecionarItemOSPeloPkid(lngITEMOSID)
  '
  Set objItemOS = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirItemOS(lngITEMOSID As Long)
  On Error GoTo trata
  Dim objItemOS As datSisMetal.clsItemOS
  Set objItemOS = New datSisMetal.clsItemOS
  
  'Excluir ItemOS
  objItemOS.ExcluirItemOS lngITEMOSID
  '
  Set objItemOS = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub CalculoQuantidadeOS(ByRef curPesoLinha As Currency, _
                               ByRef curPeso As Currency, _
                               lngLINHAID As Long, _
                               strQuantidade As String)
  On Error GoTo trata
  Dim objLinhaPerfil  As datSisMetal.clsLinhaPerfil
  Dim objRs   As ADODB.Recordset
  
  'Obter peso da linha
  Set objLinhaPerfil = New datSisMetal.clsLinhaPerfil
  Set objRs = objLinhaPerfil.SelecionarLinha(lngLINHAID)
  curPesoLinha = 0
  If Not objRs.EOF Then
    If IsNumeric(objRs.Fields("PESO_VARA").Value) Then
      curPesoLinha = objRs.Fields("PESO_VARA").Value
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  Set objLinhaPerfil = Nothing
  'CALCULO DOS PESOS
  curPeso = curPesoLinha * CCur(IIf(strQuantidade = "", 0, strQuantidade))
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub InserirItemOS(lngOSID As Long, _
                         lngLINHAID As Long, _
                         strQuantidade As String)
  On Error GoTo trata
  Dim objItemOS       As datSisMetal.clsItemOS
  '
  Set objItemOS = New datSisMetal.clsItemOS
  'Inserir ItemOS
  objItemOS.InserirItemOS lngOSID, _
                          lngLINHAID, _
                          strQuantidade
  '
  Set objItemOS = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarItemOS(lngITEMOSID As Long, _
                         lngOSID As Long, _
                         lngLINHAID As Long, _
                         strQuantidade As String)

  On Error GoTo trata
  Dim objItemOS           As datSisMetal.clsItemOS
  '
  Set objItemOS = New datSisMetal.clsItemOS
  'Atualizar ItemOS
  objItemOS.AlterarItemOS lngITEMOSID, _
                          lngOSID, _
                          lngLINHAID, _
                          strQuantidade
  '
  Set objItemOS = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

