VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarNFPeloPkid(lngNFID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objNF As datSisLoc.clsNF
  Set objNF = New datSisLoc.clsNF
  
  'Listar NF
  Set SelecionarNFPeloPkid = objNF.SelecionarNFPeloPkid(lngNFID)
  '
  Set objNF = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub InserirNF(ByRef lngPKID As Long, _
                     lngContratoId As Long, _
                     strNumero As String, _
                     strSequencial As String, _
                     strAno As String, _
                     strData As String, _
                     strDataSaida As String, _
                     strDataIniCob As String, _
                     strNroRF As String, _
                     lngOBRAID As Long)
  On Error GoTo trata
  Dim objNF     As datSisLoc.clsNF
  '
  Set objNF = New datSisLoc.clsNF
  'Inserir NF
  objNF.InserirNF lngPKID, _
                  lngContratoId, _
                  strNumero, _
                  strSequencial, _
                  strAno, _
                  strData, _
                  strDataSaida, _
                  strDataIniCob, _
                  strNroRF, _
                  lngOBRAID
  '
  Set objNF = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarNF(lngPKID As Long, _
                     lngContratoId As Long, _
                     strDataSaida As String, _
                     strDataIniCob As String, _
                     strNroRF As String, _
                     lngOBRAID As Long)
  On Error GoTo trata
  Dim objNF     As datSisLoc.clsNF
  '
  Set objNF = New datSisLoc.clsNF
  'Atualizar NF
  objNF.AlterarNF lngPKID, _
                  lngContratoId, _
                  strDataSaida, _
                  strDataIniCob, _
                  strNroRF, _
                  lngOBRAID
  '
  Set objNF = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarStatusNF(lngNFID As Long, _
                           strStatus As String)
  On Error GoTo trata
  Dim objNF     As datSisLoc.clsNF
  '
  Set objNF = New datSisLoc.clsNF
  'Atualizar Status NF
  objNF.AlterarStatusNF lngNFID, _
                        strStatus
  '
  Set objNF = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


