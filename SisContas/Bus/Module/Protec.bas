Attribute VB_Name = "mdlProtec"
Option Explicit

Public Const gsTituloDll = "Sistema de Proteção vs. 1.0"
Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
' Sleep API is declared in the form to keep the
' SetWaitableTimer code in its own re-usable module.
Private Declare Sub Sleep Lib "kernel32" ( _
    ByVal dwMilliseconds As Long)

Public Const MAX_RETRIES = 10 '50 seg

'Chaves
Global Const sSection = "System"
Global Const sKey1 = "Key_1" 'Código e quantidade de acessos
Global Const sAppName = "VB 6 Settings sctas" 'Application

Public Const strClassName As String = "mdlProtec"
Const lngTimeout As Long = 120

Global sRootPathName As String
Global sVolumeNameBuffer As String
Global lVolumeNameSize As Long
Global lVolumeSerialNumber As String '---///
Global lVolSerNum As Long '---///
Global lMaximunComponentLenght As Long
Global lFfileSystemFlags As Long
Global sFfileNameSystemBuffer As String
Global lFileSystemNameSize As Long
'
Global spBuffer As String
Global lSize As Long

Private Sub collectParams(ByRef cmd As ADODB.Command, _
                          ParamArray argparams() As Variant)
                          
  Dim params As Variant, v As Variant
  Dim i As Integer, L As Integer, u As Integer
  Dim oParam As ADODB.Parameter
  On Error GoTo trata
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
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".collectParams]", _
            Err.Description
End Sub

Function GetConnectionString(strDsn As String) As String
  On Error GoTo trata
  GetConnectionString = strDsn
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".GetConnectionString]", _
            Err.Description
End Function

Private Function CriaInstancia(ByVal pClasseObjeto As String) As Object
  Set CriaInstancia = CreateObject(pClasseObjeto)
End Function

Function RunSPReturnRS(ByVal strDsn As String, _
                       ByVal strSP As String, _
                       ParamArray params() As Variant) As ADODB.Recordset

  On Error GoTo trata
  
  ' Create the ADO objects
  Dim rs              As ADODB.Recordset
  Dim cmd             As ADODB.Command
  Dim lngErrCount     As Long
  
  Set rs = CriaInstancia("ADODB.Recordset")
  Set cmd = CriaInstancia("ADODB.Command")
  
  ' Init the ADO objects  & the stored proc parameters
  cmd.ActiveConnection = GetConnectionString(strDsn)
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
trata:
  If Err.Number = -2147467259 Then
    'ATENÇAO, OCORREU O ERRO
    lngErrCount = lngErrCount + 1
    If lngErrCount < MAX_RETRIES Then
      Sleep 5000 'sleep 5 seconds
'''      TratarErro Err.Number, _
'''                 lngErrCount & " - Erro tratado na Função Sleep", _
'''                 "[" & strClassName & "mdlGlobal.RunSPReturnRS]"
      Resume
    Else
      'Retries did not help. trate error
      Err.Raise Err.Number, _
                "[" & strClassName & "mdlGlobal.RunSPReturnRS]", _
                Err.Description
    End If
  Else
    Err.Raise Err.Number, _
              "[" & strClassName & "mdlGlobal.RunSPReturnRS]", _
              Err.Description
  End If
End Function

Function RunSP(ByVal strDsn As String, _
               ByVal strSP As String, _
               ParamArray params() As Variant)
    On Error GoTo trata

    ' Create the ADO objects
    Dim cmd             As ADODB.Command
    Dim lngErrCount     As Long

    Set cmd = CriaInstancia("ADODB.Command")

    ' Init the ADO objects & the stored proc parameters
    cmd.ActiveConnection = GetConnectionString(strDsn)
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

trata:
  If Err.Number = -2147467259 Then
    'ATENÇAO, OCORREU O ERRO
    lngErrCount = lngErrCount + 1
    If lngErrCount < MAX_RETRIES Then
      Sleep 5000 'sleep 5 seconds
'''      TratarErro Err.Number, _
'''                 lngErrCount & " - Erro tratado na Função Sleep", _
'''                 "[" & strClassName & "mdlGlobal.RunSP]"
      Resume
    Else
      'Retries did not help. trate error
      Err.Raise Err.Number, _
                "[" & strClassName & "mdlGlobal.RunSP]", _
                Err.Description
    End If
  Else
    Err.Raise Err.Number, _
              "[" & strClassName & "mdlGlobal.RunSP]", _
              Err.Description
  End If
End Function


Function Adquirir_chave(lVolumeSerialNumber As String, _
                        sDSN As String) As Boolean
  On Error GoTo trata
  Load frmRegistrarChave
  frmRegistrarChave.sVolumeKey = lVolumeSerialNumber
  frmRegistrarChave.sDataAtual = Format(Now, "DD/MM/YYYY")
  frmRegistrarChave.sBanco = sDSN
  frmRegistrarChave.Show vbModal
  Adquirir_chave = frmRegistrarChave.bRegistroValido
  Unload frmRegistrarChave
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Adquirir_chave]", _
            Err.Description
End Function


Function Retorna_Volume_Info(ByVal pVolume As Long) As String
  On Error GoTo trata
  Dim strAnexar As String
  strAnexar = "0000000000"
  strAnexar = strAnexar & CStr(pVolume)
  Retorna_Volume_Info = Right(strAnexar, 9)
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Retorna_Volume_Info]", _
            Err.Description
End Function


Public Function Gerar_Chave_HD(pSerail_HD As String) As String
  On Error GoTo trata
  Dim sRetorno As String
  '
  sRetorno = CalcDv(Mid(pSerail_HD, 1, 3))
  sRetorno = sRetorno & CalcDv(Mid(pSerail_HD, 4, 3))
  sRetorno = sRetorno & CalcDv(Mid(pSerail_HD, 7, 3))
  '
  sRetorno = sRetorno & CalcDv(Mid(pSerail_HD, 1, 6))
  sRetorno = sRetorno & CalcDv(Mid(pSerail_HD, 4, 6))
  '
  sRetorno = sRetorno & CalcDv(pSerail_HD)
  '
  Gerar_Chave_HD = sRetorno
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Gerar_Chave_HD]", _
            Err.Description
End Function

Function CalcDv(ByVal campo) As String
  On Error GoTo trata
  Dim nTot, nTam, nPos, nCont As Integer
  Dim Vet_Num() As Integer
  Dim nDiv As Integer
  Dim nUni
  Dim i
  
  
  nPos = Len(campo)
  nCont = 9
  nTam = 0
  Do While nPos > 0
    nUni = Mid$(campo, nPos, 1)
    If nCont = 9 Then
      nUni = nUni * 9
    ElseIf nCont = 8 Then
      nUni = nUni * 8
    ElseIf nCont = 7 Then
      nUni = nUni * 7
    ElseIf nCont = 6 Then
      nUni = nUni * 6
    ElseIf nCont = 5 Then
      nUni = nUni * 5
    ElseIf nCont = 4 Then
      nUni = nUni * 4
    ElseIf nCont = 3 Then
      nUni = nUni * 3
    ElseIf nCont = 2 Then
      nUni = nUni * 2
    End If
    nTam = nTam + 1
    ReDim Preserve Vet_Num(nTam)
    Vet_Num(nTam) = nUni
    nPos = nPos - 1
    nCont = nCont - 1
    If nCont < 2 Then
      nCont = 9
    End If
  Loop
  nTot = 0
  For i = 1 To nTam
    nTot = nTot + Vet_Num(i)
  Next
  nDiv = nTot Mod 11
  If nDiv <> 10 Then
    CalcDv = Format$(nDiv, "0")
  Else
    CalcDv = "0"
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".CalcDv]", _
            Err.Description
End Function

Public Function Validar_Chave(pSerail_HD As String, _
                              pChave) As Boolean
  On Error GoTo trata
  Dim bRetorno          As Boolean
  Dim sChave_Compara_1  As String
  Dim sChave_Compara_2  As String
  Dim sChave_Compara_3  As String
  Dim sChave_Compara_4  As String
  '
  bRetorno = False
  '
  sChave_Compara_1 = Mid(pSerail_HD, 1, 1) & Mid(pSerail_HD, 4, 1) & Mid(pSerail_HD, 7, 1)
  sChave_Compara_2 = Mid(pChave, 1, 1) & Mid(pChave, 4, 1) & Mid(pChave, 7, 1)
  '
  If sChave_Compara_1 = sChave_Compara_2 Then
    sChave_Compara_3 = CalcDv(Mid(pChave, 1, 2))
    sChave_Compara_3 = sChave_Compara_3 & CalcDv(Mid(pChave, 4, 2))
    sChave_Compara_3 = sChave_Compara_3 & CalcDv(Mid(pChave, 7, 2))
    '
    sChave_Compara_4 = Mid(pChave, 3, 1)
    sChave_Compara_4 = sChave_Compara_4 & Mid(pChave, 6, 1)
    sChave_Compara_4 = sChave_Compara_4 & Mid(pChave, 9, 1)
    '
    If sChave_Compara_3 = sChave_Compara_4 Then bRetorno = True
  End If
  Validar_Chave = bRetorno
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Validar_Chave]", _
            Err.Description
End Function

Public Function Validar_Chave_HD(pSerail_HD As String, _
                                 pChave) As Boolean
  On Error GoTo trata
  Dim bRetorno          As Boolean
  Dim sChave_Compara_1  As String
  Dim sChave_Compara_2  As String
  Dim sChave_Compara_3  As String
  Dim sChave_Compara_4  As String
  '
  bRetorno = False
  '
  sChave_Compara_3 = CalcDv(Mid(pSerail_HD, 1, 3))
  sChave_Compara_3 = sChave_Compara_3 & CalcDv(Mid(pSerail_HD, 4, 3))
  sChave_Compara_3 = sChave_Compara_3 & CalcDv(Mid(pSerail_HD, 7, 3))
  '
  sChave_Compara_3 = sChave_Compara_3 & CalcDv(Mid(pSerail_HD, 1, 6))
  sChave_Compara_3 = sChave_Compara_3 & CalcDv(Mid(pSerail_HD, 4, 6))
  sChave_Compara_3 = sChave_Compara_3 & CalcDv(Mid(pSerail_HD, 1, 9))
  '
  If sChave_Compara_3 = pChave Then bRetorno = True
  '
  Validar_Chave_HD = bRetorno
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Validar_Chave_HD]", _
            Err.Description
End Function


Public Function Encripta(ByVal Senha) As String
  'Propósito: criptografar a senha do usuário armazenada no banco de dados
  'Entrada: senha
  'Retorna: senha
  'caso entrada seja não criptografada a saída é criptografada e vice-versa
  On Error GoTo trata
  Dim i As Integer
  Dim str As String
  Senha = UCase$(Senha)
  For i = 1 To Len(Senha)
    str = Mid(Senha, i, 1)
    str = 255 - Asc(str)
    Encripta = Encripta & Chr(str)
  Next i
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Encripta]", _
            Err.Description
End Function

Public Sub CenterForm(frmCenter As Form)
  Dim lHeight As Integer, lWidth As Integer
  Dim lTop As Integer, lLeft As Integer
  Dim lLeftOffset As Integer, lTopOffset As Integer
  Dim ICont As Integer
  On Error GoTo trata
  lHeight = Screen.Height
  lWidth = Screen.Width
  lTop = 0
  lLeft = 0

  'Calcula "left offset"
  lLeftOffset = ((lWidth - frmCenter.Width) / 2) + lLeft
  If (lLeftOffset + frmCenter.Width > Screen.Width) Or lLeftOffset < 100 Then
    lLeftOffset = 100
  End If
  'Calcula "top offset"
  lTopOffset = ((lHeight - frmCenter.Height) / 2) + lTop
  If (lTopOffset + frmCenter.Height > Screen.Height) Or lTopOffset < 100 Then
    lTopOffset = 100
  End If
  'Centraliza o form
  frmCenter.Move lLeftOffset, lTopOffset
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".CenterForm]", _
            Err.Description
End Sub

Function Gerar_Nro_Aleatorio(ByVal pFaixa_Inicial As Long, _
                             pFaixa_Final) As Long
  On Error GoTo trata
  Randomize
  Gerar_Nro_Aleatorio = Format(CLng((pFaixa_Final - pFaixa_Inicial + 1) * Rnd + pFaixa_Inicial), "000")
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".CenterForm]", _
            Err.Description
End Function


