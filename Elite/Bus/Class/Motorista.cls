VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMotorista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarMotoristaPeloPkid(lngMOTORISTAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objMotorista As datElite.clsMotorista
  Set objMotorista = New datElite.clsMotorista
  
  'Listar Motorista
  Set SelecionarMotoristaPeloPkid = objMotorista.SelecionarMotoristaPeloPkid(lngMOTORISTAID)
  '
  Set objMotorista = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function



Public Sub InserirMotorista(lngMOTORISTAID As Long, _
                            lngBANCOID As Long, _
                            strAgencia As String, _
                            strConta As String, _
                            strExcluido As String)
  On Error GoTo trata
  Dim objMotorista     As datElite.clsMotorista
  '
  Set objMotorista = New datElite.clsMotorista
  'Inserir Motorista
  objMotorista.InserirMotorista lngMOTORISTAID, _
                                lngBANCOID, _
                                strAgencia, _
                                strConta, _
                                strExcluido
  '
  Set objMotorista = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarMotorista(lngMOTORISTAID As Long, _
                              lngBANCOID As Long, _
                              strAgencia As String, _
                              strConta As String, _
                              strExcluido As String)
  On Error GoTo trata
  Dim objMotorista     As datElite.clsMotorista
  '
  Set objMotorista = New datElite.clsMotorista
  'Atualizar Motorista
  objMotorista.AlterarMotorista lngMOTORISTAID, _
                              lngBANCOID, _
                              strAgencia, _
                              strConta, _
                              strExcluido
  '
  Set objMotorista = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

