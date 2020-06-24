VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirCLIENTE(ByVal lngCLIENTEID As Long)
  Dim clsChq    As datSisContas.clsCheque
  '
  On Error GoTo trata
  Set clsChq = New datSisContas.clsCheque
  '
  clsChq.ExcluirCLIENTE lngCLIENTEID
  '
  Set clsChq = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ExcluirCHEQUE(ByVal lngCHEQUEID As Long)
  Dim clsChq    As datSisContas.clsCheque
  '
  On Error GoTo trata
  Set clsChq = New datSisContas.clsCheque
  '
  clsChq.ExcluirCHEQUE lngCHEQUEID
  '
  Set clsChq = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirCheque(lngCLIENTEID As Long, _
                         lngBANCOID As Long, _
                         lngMOTIVODEVOLID As Long, _
                         strConta As String, _
                         strCheque As String, _
                         strAgencia As String, _
                         strValor As String, _
                         strDtDevolucao As String, _
                         strDtRecebimento As String, _
                         strDtRecuperacao As String, _
                         strStatus As String)
                         
  On Error GoTo trata
  Dim clsChq As datSisContas.clsCheque
  Set clsChq = New datSisContas.clsCheque
  
  'Inserir Venda
  clsChq.InserirCheque lngCLIENTEID, _
                       lngBANCOID, _
                       lngMOTIVODEVOLID, _
                       strConta, _
                       strCheque, _
                       strAgencia, _
                       strValor, _
                       strDtDevolucao, _
                       strDtRecebimento, _
                       strDtRecuperacao, _
                       strStatus
  '
  Set clsChq = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarCheque(lngCLIENTEID As Long, _
                         lngCHEQUEID As Long, _
                         lngBANCOID As Long, _
                         lngMOTIVODEVOLID As Long, _
                         strConta As String, _
                         strCheque As String, _
                         strAgencia As String, _
                         strValor As String, _
                         strDtDevolucao As String, _
                         strDtRecebimento As String, _
                         strDtRecuperacao As String, _
                         strStatus As String)
                         
  On Error GoTo trata
  Dim clsChq As datSisContas.clsCheque
  Set clsChq = New datSisContas.clsCheque
  
  'Inserir Venda
  clsChq.AlterarCheque lngCLIENTEID, _
                       lngCHEQUEID, _
                       lngBANCOID, _
                       lngMOTIVODEVOLID, _
                       strConta, _
                       strCheque, _
                       strAgencia, _
                       strValor, _
                       strDtDevolucao, _
                       strDtRecebimento, _
                       strDtRecuperacao, _
                       strStatus
  '
  Set clsChq = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarChequePorCPF(strCPF As String) As ADODB.Recordset
  On Error GoTo trata
  Dim clsChq As datSisContas.clsCheque
  Set clsChq = New datSisContas.clsCheque
  
  'Listar Clientes
  Set ListarChequePorCPF = clsChq.ListarChequePorCPF(strCPF)
  '
  Set clsChq = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function ListarCheque(lngCHEQUEID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsChq As datSisContas.clsCheque
  Set clsChq = New datSisContas.clsCheque
  
  'Listar Clientes
  Set ListarCheque = clsChq.ListarCheque(lngCHEQUEID)
  '
  Set clsChq = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function ListarClientePorCPF(strCPF As String, _
                                    lngCLIENTEID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsChq As datSisContas.clsCheque
  Set clsChq = New datSisContas.clsCheque
  
  'Listar Clientes
  Set ListarClientePorCPF = clsChq.ListarClientePorCPF(strCPF, lngCLIENTEID)
  '
  Set clsChq = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Sub InserirCliente(lngCLIENTEID As Long, _
                          strCPF As String, _
                          strPlaca As String, _
                          strVeiculo As String, _
                          strNome As String, _
                          strTel1 As String, _
                          strTel2 As String, _
                          strTel3 As String, _
                          strEndereco As String, _
                          strCidade As String, _
                          strBairro As String, _
                          strDtNasc As String, _
                          strEstado As String)
  On Error GoTo trata
  Dim clsChq  As datSisContas.clsCheque
  Dim objRs   As ADODB.Recordset
  '
  Set clsChq = New datSisContas.clsCheque
  
  'Inserir Venda
  clsChq.InserirCliente strCPF, _
                        strPlaca, _
                        strVeiculo, _
                        strNome, _
                        strTel1, _
                        strTel2, _
                        strTel3, _
                        strEndereco, _
                        strCidade, _
                        strBairro, _
                        strDtNasc, _
                        strEstado

  '
  Set objRs = clsChq.ListarClientePorCPF(strCPF, 0)
  If objRs.EOF Then
    lngCLIENTEID = 0
  Else
    lngCLIENTEID = objRs.Fields("PKID").Value
  End If
  '
  objRs.Close
  Set objRs = Nothing
  Set clsChq = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarCliente(lngCLIENTEID As Long, _
                          strCPF As String, _
                          strPlaca As String, _
                          strVeiculo As String, _
                          strNome As String, _
                          strTel1 As String, _
                          strTel2 As String, _
                          strTel3 As String, _
                          strEndereco As String, _
                          strCidade As String, _
                          strBairro As String, _
                          strDtNasc As String, _
                          strEstado As String)
                         
  On Error GoTo trata
  Dim clsChq As datSisContas.clsCheque
  Set clsChq = New datSisContas.clsCheque
  
  'Inserir Venda
  clsChq.AlterarCliente lngCLIENTEID, _
                          strCPF, _
                          strPlaca, _
                          strVeiculo, _
                          strNome, _
                          strTel1, _
                          strTel2, _
                          strTel3, _
                          strEndereco, _
                          strCidade, _
                          strBairro, _
                          strDtNasc, _
                          strEstado

  '
  Set clsChq = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarCliente(lngCLIENTEID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsChq As datSisContas.clsCheque
  Set clsChq = New datSisContas.clsCheque
  
  'Listar Clientes
  Set ListarCliente = clsChq.ListarCliente(lngCLIENTEID)
  '
  Set clsChq = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function


