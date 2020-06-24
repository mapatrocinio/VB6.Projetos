VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBoletoAtend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarBoletoAtendPeloPkid(lngBOLETOATENDID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objBoletoAtend As datSisMaq.clsBoletoAtend
  Set objBoletoAtend = New datSisMaq.clsBoletoAtend
  
  'Listar BoletoAtend
  Set SelecionarBoletoAtendPeloPkid = objBoletoAtend.SelecionarBoletoAtendPeloPkid(lngBOLETOATENDID)
  '
  Set objBoletoAtend = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirBoletoAtend(lngBOLETOATENDID As Long)
  On Error GoTo trata
  Dim objBoletoAtend As datSisMaq.clsBoletoAtend
  Set objBoletoAtend = New datSisMaq.clsBoletoAtend
  
  'Excluir BoletoAtend
  objBoletoAtend.ExcluirBoletoAtend lngBOLETOATENDID
  '
  Set objBoletoAtend = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirBoletoAtend(lngTURNOID As Long, _
                              lngCAIXAATENDID As Long, _
                              strNumero As String, _
                              strData As String, _
                              strStatus As String)
  On Error GoTo trata
  Dim objBoletoAtend     As datSisMaq.clsBoletoAtend
  '
  Set objBoletoAtend = New datSisMaq.clsBoletoAtend
  'Inserir BoletoAtend
  objBoletoAtend.InserirBoletoAtend lngTURNOID, _
                                    lngCAIXAATENDID, _
                                    strNumero, _
                                    strData, _
                                    strStatus
  '
  Set objBoletoAtend = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarBoletoAtend(lngPKID As Long, _
                              strNumero As String)

  On Error GoTo trata
  Dim objBoletoAtend     As datSisMaq.clsBoletoAtend
  '
  Set objBoletoAtend = New datSisMaq.clsBoletoAtend
  'Atualizar BoletoAtend
  objBoletoAtend.AlterarBoletoAtend lngPKID, _
                                    strNumero
  '
  Set objBoletoAtend = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarStatusBoletoAtend(lngPKID As Long, _
                                    strStatus As String)

  On Error GoTo trata
  Dim objBoletoAtend     As datSisMaq.clsBoletoAtend
  '
  Set objBoletoAtend = New datSisMaq.clsBoletoAtend
  'Atualizar StatusBoletoAtend
  objBoletoAtend.AlterarStatusBoletoAtend lngPKID, _
                                          strStatus
  '
  Set objBoletoAtend = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

