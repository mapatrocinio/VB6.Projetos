VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAtendimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarAtendimentoPeloPkid(lngATENDIMENTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objAtendimento As datSisMed.clsAtendimento
  Set objAtendimento = New datSisMed.clsAtendimento
  
  'Listar Atendimento
  Set SelecionarAtendimentoPeloPkid = objAtendimento.SelecionarAtendimentoPeloPkid(lngATENDIMENTOID)
  '
  Set objAtendimento = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function



Public Sub InserirAtendimento(lngGRID As Long, _
                              strData As String, _
                              strTrabalhaComScaner As String, _
                              strPathArquivo As String, _
                              strNomeArquivo As String, _
                              strDescricao As String, _
                              strDataImportacao As String, _
                              lngPRONTUARIOID As Long)
  On Error GoTo trata
  Dim objAtendimento     As datSisMed.clsAtendimento
  '
  Set objAtendimento = New datSisMed.clsAtendimento
  'Inserir Atendimento
  objAtendimento.InserirAtendimento lngGRID, _
                                    strData, _
                                    strTrabalhaComScaner, _
                                    strPathArquivo, _
                                    strNomeArquivo, _
                                    strDescricao, _
                                    strDataImportacao, _
                                    lngPRONTUARIOID
  '
  Set objAtendimento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarAtendimento(lngPKID As Long, _
                              strPathArquivo As String, _
                              strNomeArquivo As String, _
                              strDescricao As String)
  On Error GoTo trata
  Dim objAtendimento     As datSisMed.clsAtendimento
  '
  Set objAtendimento = New datSisMed.clsAtendimento
  'Atualizar Atendimento
  objAtendimento.AlterarAtendimento lngPKID, _
                                    strPathArquivo, _
                                    strNomeArquivo, _
                                    strDescricao
  '
  Set objAtendimento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub



