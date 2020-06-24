VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCartaoPromocional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub AlterarCartaoPromocional(lngCARTAOPROMOCIONALID As Long, _
                                    Numero As Long, _
                                    lngFXCARTAOPROMOCIONALID As Long)
  On Error GoTo trata
  Dim clsCartPromo As datApler.clsCartaoPromocional
  Set clsCartPromo = New datApler.clsCartaoPromocional
  
  'Alterar cartão promocional
  clsCartPromo.AlterarCartaoPromocional lngCARTAOPROMOCIONALID, _
                                        Numero, _
                                        lngFXCARTAOPROMOCIONALID, _
                                        Format(Now, "DD/MM/YYYY hh:mm")

  '
  Set clsCartPromo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirCartaoPromocional(lngCARTAOPROMOCIONALID As Long, _
                                    Numero As Long, _
                                    lngFXCARTAOPROMOCIONALID As Long)
  On Error GoTo trata
  Dim objRs As ADODB.Recordset
  Dim clsCartPromo As datApler.clsCartaoPromocional
  Set clsCartPromo = New datApler.clsCartaoPromocional
  
  'Inserir cartão promocional
  clsCartPromo.InserirCartaoPromocional Numero, _
                                        lngFXCARTAOPROMOCIONALID, _
                                        Format(Now, "DD/MM/YYYY hh:mm")

  '
  'Retornar CARTAOPROMOCIONALID
  Set objRs = clsCartPromo.ListarCartaoPromocional(Numero, 0)
  If Not objRs.EOF Then
    lngCARTAOPROMOCIONALID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set clsCartPromo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ValidaCartaoPromocional(ByVal Numero As Long, _
                                   ByVal lngCARTAOPROMOCIONALID As Long, _
                                   ByRef lngFXCARTAOPROMOCIONALID As Long, _
                                   ByRef intValidade As Integer, _
                                   ByRef lngPercDesconto As Currency, _
                                   ByRef msgErro As String)
  ' -------------------------
  ' intValidade Assume
  '   0 - Nem Chegou a validar
  '   1 - Inválida
  '   2 - Válida
  On Error GoTo trata
  Dim clsCartPromo  As datApler.clsCartaoPromocional
  Dim objRs         As ADODB.Recordset
  '
  Set clsCartPromo = New datApler.clsCartaoPromocional
  
  msgErro = ""
  intValidade = 0
  'Validar se cartão promocional já lançado
  Set objRs = clsCartPromo.ListarCartaoPromocional(Numero, _
                                                   lngCARTAOPROMOCIONALID)
  '
  If Not objRs.EOF Then
    'msgErro = "cartão promocional já cadastrado."
  End If
  objRs.Close
  Set objRs = Nothing
  If msgErro = "" Then
    'Validar fx cartão promocional
    Set objRs = clsCartPromo.ListarFxCartaoPromocionalPelaFaixa(Numero)
    '
    If objRs.EOF Then
      msgErro = "Cartão promocional informado não está dentro de uma faixa válida."
    Else
      'FAIXA CADASTRADA
      'PEGA lngFXCARTAOPROMOCIONALID
      lngFXCARTAOPROMOCIONALID = objRs.Fields("PKID").Value
      lngPercDesconto = objRs.Fields("PERCDESCONTO").Value
      'VALIDA DATA DE VALIDADE DO CARTÃO
      intValidade = 1 'fora da validade
      If IsDate(objRs.Fields("VALIDADE").Value) Then 'data cadastrada
        If (CDate(Format(objRs.Fields("VALIDADE").Value, "YYYY/MM/DD") & " 23:59")) < Now Then
          'Fora Data de validade
          msgErro = "Data cadastrada na faixa de cartão fora da data de validade."
        Else
          intValidade = 2
          
        End If
      Else
        msgErro = "Data cadastrada na faixa de cartão inválida."
      End If
      
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  '
  Set clsCartPromo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarCartaoPromocionalPeloID(lngCARTAOPROMOCIONALID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsCartPromo As datApler.clsCartaoPromocional
  Set clsCartPromo = New datApler.clsCartaoPromocional

  'Listar Cartão Promocional
  Set ListarCartaoPromocionalPeloID = clsCartPromo.ListarCartaoPromocionalPeloID(lngCARTAOPROMOCIONALID)
  '
  Set clsCartPromo = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub AlterarFxCartaoPromocional(lngFXCARTAOPROMOCIONALID As Long, _
                                      FxInicial As Long, _
                                      FxFinal As Long, _
                                      PercDesc As String, _
                                      lngCLIENTECARTAOPROMOCIONALID As Long, _
                                      Validade As String)
  On Error GoTo trata
  Dim clsCartPromo As datApler.clsCartaoPromocional
  Set clsCartPromo = New datApler.clsCartaoPromocional
  
  'Alterar Fx cartão promocional
  clsCartPromo.AlterarFxCartaoPromocional lngFXCARTAOPROMOCIONALID, _
                                          FxInicial, _
                                          FxFinal, _
                                          PercDesc, _
                                          lngCLIENTECARTAOPROMOCIONALID, _
                                          Validade

  '
  Set clsCartPromo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarClienteCartaoPromocionalPeloNome(strCliente As String) As ADODB.Recordset
  On Error GoTo trata
  Dim clsCartPromo As datApler.clsCartaoPromocional
  Set clsCartPromo = New datApler.clsCartaoPromocional
  
  'Listar Cliente Cartão Promocional Pelo nome
  Set ListarClienteCartaoPromocionalPeloNome = clsCartPromo.ListarClienteCartaoPromocionalPeloNome(strCliente)
  '
  Set clsCartPromo = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub InserirClienteCartaoPromocional(strNome As String, _
                                           strNumero As String)

  On Error GoTo trata
  Dim clsCartPromo  As datApler.clsCartaoPromocional
  '
  Set clsCartPromo = New datApler.clsCartaoPromocional
  
  'Inserir Cliente Cartão promocional
  clsCartPromo.InserirClienteCartaoPromocional strNome, _
                                               strNumero
  '
  Set clsCartPromo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarClienteCartaoPromocional(lngCLIENTECARTAOPROMOCIONALID As Long, _
                                           strNome As String, _
                                           strNumero As String)
                         
  On Error GoTo trata
  Dim clsCartPromo As datApler.clsCartaoPromocional
  Set clsCartPromo = New datApler.clsCartaoPromocional
  
  'Inserir Cliente Cartão promocional
  clsCartPromo.AlterarClienteCartaoPromocional lngCLIENTECARTAOPROMOCIONALID, _
                                               strNome, _
                                               strNumero

  '
  Set clsCartPromo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarClienteCartaoPromocional(lngCLIENTECARTAOPROMOCIONALID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsCartPromo As datApler.clsCartaoPromocional
  Set clsCartPromo = New datApler.clsCartaoPromocional
  
  'Listar Cartão Promocional
  Set ListarClienteCartaoPromocional = clsCartPromo.ListarClienteCartaoPromocional(lngCLIENTECARTAOPROMOCIONALID)
  '
  Set clsCartPromo = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirClienteCartaoPromocional(lngCLIENTECARTAOPROMOCIONALID As Long)

  On Error GoTo trata
  Dim clsCartPromo As datApler.clsCartaoPromocional
  Set clsCartPromo = New datApler.clsCartaoPromocional
  
  'Excluir cliente cartão promocional
  clsCartPromo.ExlcuirClienteCartaoPromocional lngCLIENTECARTAOPROMOCIONALID
  '
  Set clsCartPromo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarFxCartaoPromocional(lngFXCARTAOPROMOCIONALID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsCartPromo As datApler.clsCartaoPromocional
  Set clsCartPromo = New datApler.clsCartaoPromocional
  
  'Listar Faixa Cartão Promocional
  Set ListarFxCartaoPromocional = clsCartPromo.ListarFxCartaoPromocional(lngFXCARTAOPROMOCIONALID)
  '
  Set clsCartPromo = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ValidaFxCartaoPromocional(NumeroInicial As Long, _
                                     NumeroFinal As Long, _
                                     lngFXCARTAOPROMOCIONALID As Long, _
                                     msgErro As String)

  On Error GoTo trata
  Dim clsCartPromo  As datApler.clsCartaoPromocional
  Dim objRs         As ADODB.Recordset
  '
  Set clsCartPromo = New datApler.clsCartaoPromocional
  
  'Validar fx cartão promocional
  Set objRs = clsCartPromo.ValidaFxCartaoPromocional(NumeroInicial, _
                                                     NumeroFinal, _
                                                     lngFXCARTAOPROMOCIONALID)
  '
  msgErro = ""
  If Not objRs.EOF Then
    If objRs.RecordCount > 0 Then
      msgErro = "Faixa de cartão promocional informada está dentro de uma faixa já cadastrada."
    End If
  End If
  Set clsCartPromo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirFxCartaoPromocional(FxInicial As Long, _
                                      FxFinal As Long, _
                                      PercDesc As String, _
                                      lngCLIENTECARTAOPROMOCIONALID As Long, _
                                      Validade As String)
  On Error GoTo trata
  Dim clsCartPromo As datApler.clsCartaoPromocional
  Set clsCartPromo = New datApler.clsCartaoPromocional
  
  'Inserir cartão promocional
  clsCartPromo.InserirFxCartaoPromocional FxInicial, _
                                          FxFinal, _
                                          PercDesc, _
                                          lngCLIENTECARTAOPROMOCIONALID, _
                                          Validade

  '
  Set clsCartPromo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub ExcluirCartaoPromocional(lngCARTAOPROMOCIONALID As Long)

  On Error GoTo trata
  Dim clsCartPromo As datApler.clsCartaoPromocional
  Set clsCartPromo = New datApler.clsCartaoPromocional
  
  'Excluir cartão promocional
  clsCartPromo.ExlcuirCartaoPromocional lngCARTAOPROMOCIONALID
  '
  Set clsCartPromo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ExcluirFxCartaoPromocional(lngFAIXACARTAOPROMOCIONALID As Long)

  On Error GoTo trata
  Dim clsCartPromo As datApler.clsCartaoPromocional
  Set clsCartPromo = New datApler.clsCartaoPromocional
  
  'Excluir cartão promocional
  clsCartPromo.ExlcuirFxCartaoPromocional lngFAIXACARTAOPROMOCIONALID
  '
  Set clsCartPromo = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

