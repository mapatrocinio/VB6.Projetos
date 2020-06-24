Attribute VB_Name = "Database"
Option Explicit
Option Compare Text

Const lngTimeout As Long = 120
Public Const strClassName As String = "DatApler"

Function GetConnectionString() As String
  On Error GoTo errorHandler
  Dim objRegistro       As datSisLoc.clsRegistro
  Dim strNomeBD         As String
  Set objRegistro = New datSisLoc.clsRegistro
  strNomeBD = objRegistro.RetornarChaveRegistro(TITULOSISTEMA, _
                                                "ServidorBD")
  
  Set objRegistro = Nothing
  GetConnectionString = "Provider=SQLOLEDB" & _
    ";Initial Catalog=SisLoc" & _
    ";Data Source=" & strNomeBD & _
    ";User Id=SA" & _
    ";Password=21321"
  'GetConnectionString = "DSN=SisLoc" & _
    ";Password=" & psSenhaPerm
  'GetConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source= c:\Program Files\Microsoft Office\" & _
      "Office\Samples\Northwind.mdb;"
       
  Exit Function
errorHandler:
  Err.Raise Err.Number, _
            Err.Source & ".[Database.GetConnectionString]", _
            Err.Description
End Function


Function RunSPReturnRS(ByVal strSP As String, _
                       ParamArray params() As Variant) As ADODB.Recordset

  On Error GoTo errorHandler
  
  ' Create the ADO objects
  Dim rs  As ADODB.Recordset
  Dim cmd As ADODB.Command
  Dim lngErrCount     As Long

  lngErrCount = 0
  
  Set rs = CriaInstancia("ADODB.Recordset")
  Set cmd = CriaInstancia("ADODB.Command")
  
  ' Init the ADO objects  & the stored proc parameters
  cmd.ActiveConnection = GetConnectionString
  cmd.CommandTimeout = lngTimeout
  cmd.CommandText = strSP
  cmd.CommandType = adCmdText
  
  collectParams cmd, params
  
  ' Execute the query for readonly
  rs.CursorLocation = adUseClient
  rs.Open cmd, , adOpenForwardOnly, adLockReadOnly
  
  ' Disconnect the recordset
  Set cmd.ActiveConnection = Nothing
  Set cmd = Nothing
  Set rs.ActiveConnection = Nothing
  
  ' Return the resultant recordset
  Set RunSPReturnRS = rs

  Exit Function
errorHandler:
  Err.Raise Err.Number, _
            "[" & strClassName & "Database.RunSPReturnRS]" & _
            Err.Description & " | sql: " & strSP
End Function

Function RunSP(ByVal strSP As String, ParamArray params() As Variant)
  On Error GoTo errorHandler

  ' Create the ADO objects
  Dim cmd             As ADODB.Command
  Dim lngErrCount     As Long
  lngErrCount = 0
  Set cmd = CriaInstancia("ADODB.Command")
  

  ' Init the ADO objects & the stored proc parameters
  cmd.ActiveConnection = GetConnectionString
  cmd.CommandTimeout = lngTimeout
  cmd.CommandText = strSP
  cmd.CommandType = adCmdText
  collectParams cmd, params

  ' Execute the query without returning a recordset
  cmd.Execute , , adExecuteNoRecords
  ' Disconnect the recordset and clean up
  Set cmd.ActiveConnection = Nothing
  Set cmd = Nothing

  Exit Function

errorHandler:
'  If Err.Number = -2147467259 Then
'    'ATENÇAO, OCORREU O ERRO
'    lngErrCount = lngErrCount + 1
'    If lngErrCount < MAX_RETRIES Then
'      Sleep 5000 'sleep 5 seconds
'      TratarErro Err.Number, _
'                 lngErrCount & " - Erro tratado na Função Sleep", _
'                 "[" & strClassName & "Database.RunSP]"
'      Resume
'    Else
'      'Retries did not help. trate error
'      Err.Raise Err.Number, _
'                "[" & strClassName & "Database.RunSP]", _
'                Err.Description & " | sql: " & strSP
'    End If
'  Else
    Err.Raise Err.Number, _
              "[" & strClassName & "Database.RunSP]", _
              Err.Description & " | sql: " & strSP
'  End If
End Function

Private Function CriaInstancia(ByVal pClasseObjeto As String) As Object
  Set CriaInstancia = CreateObject(pClasseObjeto)
End Function

Private Sub collectParams(ByRef cmd As ADODB.Command, _
                          ParamArray argparams() As Variant)
                          
  Dim params As Variant, v As Variant
  Dim i As Integer, L As Integer, u As Integer
  Dim oParam As ADODB.Parameter
On Error GoTo errorHandler
  params = argparams(0)
  For i = LBound(params) To UBound(params)
    L = LBound(params(i))
    u = UBound(params(i))
    ' Check for nulls.
    
    If VarType(params(i)(3)) = vbString Then
      v = IIf(params(i)(3) = "", Null, params(i)(3))
    Else
      v = params(i)(3)
    End If
    Set oParam = cmd.CreateParameter(params(i)(0), params(i)(1), adParamInput, params(i)(2), v)
    If oParam.Type = adNumeric Then
      oParam.Precision = params(i)(4)
      oParam.NumericScale = params(i)(5)
    End If
    cmd.Parameters.Append oParam
  Next i
  Exit Sub
errorHandler:
  Err.Raise Err.Number, _
            Err.Source & ".[Database.collectParams]", _
            Err.Description
End Sub


