VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBoletoArrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarBoletoArrecPeloPkid(lngBOLETOARRECID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objBoletoArrec As datSisMaq.clsBoletoArrec
  Set objBoletoArrec = New datSisMaq.clsBoletoArrec
  
  'Listar BoletoArrec
  Set SelecionarBoletoArrecPeloPkid = objBoletoArrec.SelecionarBoletoArrecPeloPkid(lngBOLETOARRECID)
  '
  Set objBoletoArrec = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirBoletoArrec(lngBOLETOARRECID As Long)
  On Error GoTo trata
  Dim objBoletoArrec As datSisMaq.clsBoletoArrec
  Set objBoletoArrec = New datSisMaq.clsBoletoArrec
  
  'Excluir BoletoArrec
  objBoletoArrec.ExcluirBoletoArrec lngBOLETOARRECID
  '
  Set objBoletoArrec = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirBoletoArrec(lngTURNOID As Long, _
                              lngCAIXAARRECID As Long, _
                              strNumero As String, _
                              strData As String, _
                              strStatus As String)
  On Error GoTo trata
  Dim objBoletoArrec     As datSisMaq.clsBoletoArrec
  '
  Set objBoletoArrec = New datSisMaq.clsBoletoArrec
  'Inserir BoletoArrec
  objBoletoArrec.InserirBoletoArrec lngTURNOID, _
                                    lngCAIXAARRECID, _
                                    strNumero, _
                                    strData, _
                                    strStatus
  '
  Set objBoletoArrec = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarBoletoArrec(lngPKID As Long, _
                              strNumero As String)

  On Error GoTo trata
  Dim objBoletoArrec     As datSisMaq.clsBoletoArrec
  '
  Set objBoletoArrec = New datSisMaq.clsBoletoArrec
  'Atualizar BoletoArrec
  objBoletoArrec.AlterarBoletoArrec lngPKID, _
                                    strNumero
  '
  Set objBoletoArrec = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarStatusBoletoArrec(lngPKID As Long, _
                                    strStatus As String)

  On Error GoTo trata
  Dim objBoletoArrec     As datSisMaq.clsBoletoArrec
  '
  Set objBoletoArrec = New datSisMaq.clsBoletoArrec
  'Atualizar StatusBoletoArrec
  objBoletoArrec.AlterarStatusBoletoArrec lngPKID, _
                                          strStatus
  '
  Set objBoletoArrec = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

