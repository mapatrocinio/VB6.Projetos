VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarEntradaPeloPkid(lngENTRADAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objEntrada As datSisMaq.clsEntrada
  Set objEntrada = New datSisMaq.clsEntrada
  
  'Listar Entrada
  Set SelecionarEntradaPeloPkid = objEntrada.SelecionarEntradaPeloPkid(lngENTRADAID)
  '
  Set objEntrada = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirEntrada(lngENTRADAID As Long)
  On Error GoTo trata
  Dim objEntrada As datSisMaq.clsEntrada
  Set objEntrada = New datSisMaq.clsEntrada
  
  'Excluir Entrada
  objEntrada.ExcluirEntrada lngENTRADAID
  '
  Set objEntrada = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirEntrada(lngTURNOID As Long, _
                          strValor As String, _
                          strStatus As String, _
                          strData As String, _
                          lngGERENTEID As Long)
  On Error GoTo trata
  Dim objEntrada     As datSisMaq.clsEntrada
  '
  Set objEntrada = New datSisMaq.clsEntrada
  'Inserir Entrada
  objEntrada.InserirEntrada lngTURNOID, _
                            strValor, _
                            strStatus, _
                            strData, _
                            lngGERENTEID
  '
  Set objEntrada = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarEntrada(lngPKID As Long, _
                          strValor As String)

  On Error GoTo trata
  Dim objEntrada     As datSisMaq.clsEntrada
  '
  Set objEntrada = New datSisMaq.clsEntrada
  'Atualizar Entrada
  objEntrada.AlterarEntrada lngPKID, _
                            strValor
  '
  Set objEntrada = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

