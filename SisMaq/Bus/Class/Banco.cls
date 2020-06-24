VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirBanco(strNome As String, _
                        strNumero As String)

  On Error GoTo trata
  Dim objBanco As datSisMaq.clsBanco
  '
  Set objBanco = New datSisMaq.clsBanco
  
  'Inserir Banco
  objBanco.InserirBanco strNome, _
                        strNumero
  '
  Set objBanco = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarBanco(lngBANCOID As Long, _
                        strNome As String, _
                        strNumero As String)
                         
  On Error GoTo trata
  Dim objBanco As datSisMaq.clsBanco
  Set objBanco = New datSisMaq.clsBanco
  
  'Inserir Banco
  objBanco.AlterarBanco lngBANCOID, _
                        strNome, _
                        strNumero

  '
  Set objBanco = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarBanco(lngBANCOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objBanco As datSisMaq.clsBanco
  Set objBanco = New datSisMaq.clsBanco
  
  'Listar Banco
  Set ListarBanco = objBanco.ListarBanco(lngBANCOID)
  '
  Set objBanco = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirBanco(lngBANCOID As Long)

  On Error GoTo trata
  Dim objBanco As datSisMaq.clsBanco
  Set objBanco = New datSisMaq.clsBanco
  
  'Excluir Banco
  objBanco.ExlcuirBanco lngBANCOID
  '
  Set objBanco = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

