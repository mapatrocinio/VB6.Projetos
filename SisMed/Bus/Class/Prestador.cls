VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrestador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarPrestadorPeloPkid(lngPRESTADORID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objPrestador As datSisMed.clsPrestador
  Set objPrestador = New datSisMed.clsPrestador
  
  'Listar Prestador
  Set SelecionarPrestadorPeloPkid = objPrestador.SelecionarPrestadorPeloPkid(lngPRESTADORID)
  '
  Set objPrestador = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function



Public Sub InserirPrestador(lngPRONTUARIOID As Long, _
                            lngFUNCAOID As Long, _
                            strCRM As String, _
                            blnTrabalhaComScaner As Boolean, _
                            strIndAceitaCheque As String, _
                            strExcluido As String)
  On Error GoTo trata
  Dim objPrestador     As datSisMed.clsPrestador
  '
  Set objPrestador = New datSisMed.clsPrestador
  'Inserir Prestador
  objPrestador.InserirPrestador lngPRONTUARIOID, _
                            lngFUNCAOID, _
                            strCRM, _
                            blnTrabalhaComScaner, _
                            strIndAceitaCheque, _
                            strExcluido
  '
  Set objPrestador = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarPrestador(lngPRONTUARIOID As Long, _
                            lngFUNCAOID As Long, _
                            strCRM As String, _
                            blnTrabalhaComScaner As Boolean, _
                            strIndAceitaCheque As String, _
                            strExcluido As String, _
                            strDataDesativa As String)
  On Error GoTo trata
  Dim objPrestador     As datSisMed.clsPrestador
  '
  Set objPrestador = New datSisMed.clsPrestador
  'Atualizar Prestador
  objPrestador.AlterarPrestador lngPRONTUARIOID, _
                            lngFUNCAOID, _
                            strCRM, _
                            blnTrabalhaComScaner, _
                            strIndAceitaCheque, _
                            strExcluido, _
                            strDataDesativa
  '
  Set objPrestador = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub



