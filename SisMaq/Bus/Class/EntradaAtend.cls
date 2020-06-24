VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEntradaAtend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarEntradaAtendPeloPkid(lngENTRADAATENDID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objEntradaAtend As datSisMaq.clsEntradaAtend
  Set objEntradaAtend = New datSisMaq.clsEntradaAtend
  
  'Listar EntradaAtend
  Set SelecionarEntradaAtendPeloPkid = objEntradaAtend.SelecionarEntradaAtendPeloPkid(lngENTRADAATENDID)
  '
  Set objEntradaAtend = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirEntradaAtend(lngENTRADAATENDID As Long)
  On Error GoTo trata
  Dim objEntradaAtend As datSisMaq.clsEntradaAtend
  Set objEntradaAtend = New datSisMaq.clsEntradaAtend
  
  'Excluir EntradaAtend
  objEntradaAtend.ExcluirEntradaAtend lngENTRADAATENDID
  '
  Set objEntradaAtend = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirEntradaAtend(lngTURNOID As Long, _
                               lngCAIXAATENDID As Long, _
                               strValor As String, _
                               strData As String)
  On Error GoTo trata
  Dim objEntradaAtend     As datSisMaq.clsEntradaAtend
  '
  Set objEntradaAtend = New datSisMaq.clsEntradaAtend
  'Inserir EntradaAtend
  objEntradaAtend.InserirEntradaAtend lngTURNOID, _
                                      lngCAIXAATENDID, _
                                      strValor, _
                                      strData
  '
  Set objEntradaAtend = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarEntradaAtend(lngPKID As Long, _
                               strValor As String)

  On Error GoTo trata
  Dim objEntradaAtend     As datSisMaq.clsEntradaAtend
  '
  Set objEntradaAtend = New datSisMaq.clsEntradaAtend
  'Atualizar EntradaAtend
  objEntradaAtend.AlterarEntradaAtend lngPKID, _
                                      strValor
  '
  Set objEntradaAtend = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

