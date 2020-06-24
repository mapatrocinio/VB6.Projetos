VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaquina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarEquipamentoPeloNumero(strNumero As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objMaquina As datSisMaq.clsMaquina
  Set objMaquina = New datSisMaq.clsMaquina
  
  'Listar Equipamento
  Set SelecionarEquipamentoPeloNumero = objMaquina.SelecionarEquipamentoPeloNumero(strNumero)
  '
  Set objMaquina = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarEquipamentoPeloPkid(lngEQUIPAMENTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objMaquina As datSisMaq.clsMaquina
  Set objMaquina = New datSisMaq.clsMaquina
  
  'Listar Equipamento
  Set SelecionarEquipamentoPeloPkid = objMaquina.SelecionarEquipamentoPeloPkid(lngEQUIPAMENTOID)
  '
  Set objMaquina = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirEquipamento(lngEQUIPAMENTOID As Long)
  On Error GoTo trata
  Dim objMaquina As datSisMaq.clsMaquina
  Set objMaquina = New datSisMaq.clsMaquina
  
  'Excluir Equipamento
  objMaquina.ExcluirEquipamento lngEQUIPAMENTOID
  '
  Set objMaquina = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirMaquina(lngEQUIPAMENTOID As Long, _
                          lngTIPOID As Long, _
                          strInicio As String, _
                          strStatus As String)
  On Error GoTo trata
  Dim objMaquina     As datSisMaq.clsMaquina
  '
  Set objMaquina = New datSisMaq.clsMaquina
  'Inserir Máquina
  objMaquina.InserirMaquina lngEQUIPAMENTOID, _
                            lngTIPOID, _
                            strInicio, _
                            strStatus
  '
  Set objMaquina = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirEquipamento(ByRef lngPKID As Long, _
                              lngSERIEID As Long, _
                              strNumero As String, _
                              strCoeficiente As String, _
                              strStatus As String, _
                              lngTIPOID As Long, _
                              strDataInicio As String, _
                              strStatusMaq As String)
  On Error GoTo trata
  Dim objMaquina     As datSisMaq.clsMaquina
  '
  Set objMaquina = New datSisMaq.clsMaquina
  'Inserir Equipamento
  objMaquina.InserirEquipamento lngPKID, _
                                lngSERIEID, _
                                strNumero, _
                                strCoeficiente, _
                                strStatus
  'Inserir Máquina
  objMaquina.InserirMaquina lngPKID, _
                            lngTIPOID, _
                            strDataInicio, _
                            strStatusMaq
  '
  Set objMaquina = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub



Public Sub AlterarMaquina(lngPKID As Long, _
                          lngTIPOID As Long, _
                          strDataTermino As String, _
                          strStatus As String, _
                          Optional strUsuario As String)

  On Error GoTo trata
  Dim objMaquina     As datSisMaq.clsMaquina
  '
  Set objMaquina = New datSisMaq.clsMaquina
  'Atualizar Maquina
  objMaquina.AlterarMaquina lngPKID, _
                            lngTIPOID, _
                            strDataTermino, _
                            strStatus, _
                            strUsuario
  '
  Set objMaquina = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarEquipamento(lngPKID As Long, _
                              lngSERIEID As Long, _
                              strNumero As String, _
                              strCoeficiente As String, _
                              strStatus As String)

  On Error GoTo trata
  Dim objMaquina     As datSisMaq.clsMaquina
  '
  Set objMaquina = New datSisMaq.clsMaquina
  'Atualizar Equipamento
  objMaquina.AlterarEquipamento lngPKID, _
                                lngSERIEID, _
                                strNumero, _
                                strCoeficiente, _
                                strStatus
  '
  Set objMaquina = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

