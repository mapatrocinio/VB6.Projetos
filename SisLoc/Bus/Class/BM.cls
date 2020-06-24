VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarBM(lngBMId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objBM As datSisLoc.clsBM
  Set objBM = New datSisLoc.clsBM
  
  'Selecionar BM
  Set SelecionarBM = objBM.SelecionarBM(lngBMId)
  '
  Set objBM = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirBM(lngBMId As Long)
  On Error GoTo trata
  Dim objBM As datSisLoc.clsBM
  Set objBM = New datSisLoc.clsBM
  
  'Excluir BM
  objBM.ExcluirBM lngBMId
  '
  Set objBM = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirBM(lngContratoId As Long, _
                     strNumero As String, _
                     strMedicao As String, _
                     strDataEmissao As String, _
                     strDataInicio As String, _
                     strDataTermino As String)
  On Error GoTo trata
  Dim objBM     As datSisLoc.clsBM
  '
  Set objBM = New datSisLoc.clsBM
  'Inserir BM
  objBM.InserirBM lngContratoId, _
                  strNumero, _
                  strMedicao, _
                  strDataEmissao, _
                  strDataInicio, _
                  strDataTermino
  '
  Set objBM = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarBM(lngPKID As Long, _
                     lngContratoId As Long, _
                     strMedicao As String, _
                     strDataInicio As String, _
                     strDataTermino As String)
  On Error GoTo trata
  Dim objBM     As datSisLoc.clsBM
  '
  Set objBM = New datSisLoc.clsBM
  'Atualizar BM
  objBM.AlterarBM lngPKID, _
                  lngContratoId, _
                  strMedicao, _
                  strDataInicio, _
                  strDataTermino
  '
  Set objBM = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

