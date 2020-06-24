VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpTroca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarEmpTrocaPeloPkid(lngEMPTROCAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objEmpTroca As datSisMaq.clsEmpTroca
  Set objEmpTroca = New datSisMaq.clsEmpTroca
  
  'Listar EmpTroca
  Set SelecionarEmpTrocaPeloPkid = objEmpTroca.SelecionarEmpTrocaPeloPkid(lngEMPTROCAID)
  '
  Set objEmpTroca = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirEmpTroca(lngEMPTROCAID As Long)
  On Error GoTo trata
  Dim objEmpTroca As datSisMaq.clsEmpTroca
  Set objEmpTroca = New datSisMaq.clsEmpTroca
  
  'Excluir EmpTroca
  objEmpTroca.ExcluirEmpTroca lngEMPTROCAID
  '
  Set objEmpTroca = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirEmpTroca(lngTURNOID As Long, _
                           lngTIPOPGTOID As Long, _
                           strValor As String, _
                           strData As String, _
                           strNome As String)
  On Error GoTo trata
  Dim objEmpTroca     As datSisMaq.clsEmpTroca
  '
  Set objEmpTroca = New datSisMaq.clsEmpTroca
  'Inserir EmpTroca
  objEmpTroca.InserirEmpTroca lngTURNOID, _
                              lngTIPOPGTOID, _
                              strValor, _
                              strData, _
                              strNome
  '
  Set objEmpTroca = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarEmpTroca(lngPKID As Long, _
                           lngTIPOPGTOID As Long, _
                           strValor As String, _
                           strNome As String)

  On Error GoTo trata
  Dim objEmpTroca     As datSisMaq.clsEmpTroca
  '
  Set objEmpTroca = New datSisMaq.clsEmpTroca
  'Atualizar EmpTroca
  objEmpTroca.AlterarEmpTroca lngPKID, _
                              lngTIPOPGTOID, _
                              strValor, _
                              strNome
  '
  Set objEmpTroca = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

