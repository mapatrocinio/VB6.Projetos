VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProtec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const gsTituloDll = "Sistema de Proteção vs. 1.0"


Public Function Valida_Data(sData As String) As Boolean
  On Error GoTo trata
  Dim sChave2         As String
  Dim sDataIni        As String
  Dim bRetorno        As Boolean
  Dim sDv             As String
  Dim sDv_Calc_1      As String
  '
  bRetorno = False
  '
  If Len(Trim(sData)) <> 14 Then
    MsgBox "Erro C_2.1 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
    Valida_Data = bRetorno
    Exit Function
  End If
  '
  sChave2 = Encripta(sData)
  sDataIni = Mid(sChave2, 1, 2) & "/" & Mid(sChave2, 4, 2) & "/" & Mid(sChave2, 7, 4)
  sDv = Mid(sChave2, 3, 1) & Mid(sChave2, 6, 1) & Mid(sChave2, 11, 4)
  sDv_Calc_1 = CalcDv(Mid(sChave2, 1, 2))
  sDv_Calc_1 = sDv_Calc_1 & CalcDv(Mid(sChave2, 4, 2))
  sDv_Calc_1 = sDv_Calc_1 & CalcDv(Mid(sChave2, 7, 4))
  sDv_Calc_1 = sDv_Calc_1 & CalcDv(Mid(sChave2, 1, 10))
  sDv_Calc_1 = sDv_Calc_1 & CalcDv(Mid(sChave2, 1, 12))
  sDv_Calc_1 = sDv_Calc_1 & CalcDv(Mid(sChave2, 1, 3))
  '
  If Not IsDate(sDataIni) Then
    MsgBox "Erro C_2.2 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
    Valida_Data = bRetorno
    Exit Function
  End If
  If sDv <> sDv_Calc_1 Then
    MsgBox "Erro C_2.3 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
    Valida_Data = bRetorno
    Exit Function
  End If
  bRetorno = True
  '
  Valida_Data = bRetorno
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Valida_Data]", _
            Err.Description
End Function

Public Function Valida_Data_Atual(sData As String) As Boolean
  On Error GoTo trata
  Dim sChave2       As String
  Dim sDataIni      As String
  Dim bRetorno      As Boolean
  Dim sDv           As String
  Dim sDv_Calc_1    As String
  '
  bRetorno = False
  '
  If Len(Trim(sData)) <> 21 Then
    MsgBox "Erro C_3.1 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
    Valida_Data_Atual = bRetorno
    Exit Function
  End If
  '
  sChave2 = Encripta(sData)
  sDataIni = Mid(sChave2, 1, 2) & "/" & Mid(sChave2, 4, 2) & "/" & Mid(sChave2, 7, 4) & " " & Mid(sChave2, 15, 2) & ":" & Mid(sChave2, 17, 2) & ":" & Mid(sChave2, 19, 2)
  sDv = Mid(sChave2, 3, 1) & Mid(sChave2, 6, 1) & Mid(sChave2, 11, 4) & Mid(sChave2, 21, 1)
  sDv_Calc_1 = CalcDv(Mid(sChave2, 1, 2))
  sDv_Calc_1 = sDv_Calc_1 & CalcDv(Mid(sChave2, 4, 2))
  sDv_Calc_1 = sDv_Calc_1 & CalcDv(Mid(sChave2, 7, 4))
  sDv_Calc_1 = sDv_Calc_1 & CalcDv(Mid(sChave2, 1, 11))
  sDv_Calc_1 = sDv_Calc_1 & CalcDv(Mid(sChave2, 1, 6))
  sDv_Calc_1 = sDv_Calc_1 & CalcDv(Mid(sChave2, 1, 13))
  sDv_Calc_1 = sDv_Calc_1 & CalcDv(Mid(sChave2, 15, 6))
  '
  If Not IsDate(sDataIni) Then
    MsgBox "Erro C_3.2 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
    Valida_Data_Atual = bRetorno
    Exit Function
  End If
  If sDv <> sDv_Calc_1 Then
    MsgBox "Erro C_3.3 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
    Valida_Data_Atual = bRetorno
    Exit Function
  End If
  bRetorno = True
  '
  Valida_Data_Atual = bRetorno
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Valida_Data_Atual]", _
            Err.Description

End Function

Public Function Valida_Chaves_ok(pDsn As String, _
                                 pSerail_HD As String) As Boolean
  Dim QtdValidos        As Integer
  Dim QtdInvalidos      As Integer
  Dim sQtd_Vez          As String
  Dim bRetorno          As Boolean
  Dim rs                As ADODB.Recordset
  Dim sSql              As String
  Dim objGeral          As busSisMed.clsGeral
  '
  Set objGeral = New busSisMed.clsGeral
  bRetorno = False
  '
  sSql = "Select * from Chave;"
  Set rs = objGeral.ExecutarSQL(sSql)
  QtdValidos = 0
  QtdInvalidos = 0
  Do While Not rs.EOF
    'Verifica se todas as chaves são válidas no bd
    '
    If Not Validar_Chave(pSerail_HD, Encripta(rs!Chave1)) Then
      QtdInvalidos = QtdInvalidos + 1
    Else
      QtdValidos = QtdValidos + 1
    End If
    '
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  Set objGeral = Nothing
  '
  If QtdValidos = 0 Then
    MsgBox "Erro C_6.1 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
    Valida_Chaves_ok = bRetorno
    Exit Function
  ElseIf QtdValidos > 1 Then
    MsgBox "Erro C_6.2 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
    Valida_Chaves_ok = bRetorno
    Exit Function
  End If
  bRetorno = True
  Valida_Chaves_ok = bRetorno
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Valida_Chaves_Ok]", _
            Err.Description
End Function

Public Function Valida_Chaves_Duplic(pDsn As String) As Boolean
  On Error GoTo trata
  Dim sQtd_Vez    As String
  Dim bRetorno    As Boolean
  Dim rs          As ADODB.Recordset
  Dim rsDuplic    As ADODB.Recordset
  Dim sSql        As String
  Dim objGeral    As busSisMed.clsGeral
  '
  Set objGeral = New busSisMed.clsGeral
  '
  bRetorno = False
  '
  sSql = "Select * from Chave;"
  Set rs = objGeral.ExecutarSQL(sSql)
  Do While Not rs.EOF
    'Verifica se a chave não está duplicada
    '
    sSql = "Select Count(*) from Chave Where Chave1 = " & Formata_Dados(rs!Chave1, tpDados_Texto, tpNulo_Aceita) & " And PKID <> " & rs!PKID & ";"
    Set rsDuplic = objGeral.ExecutarSQL(sSql)
    If rsDuplic(0) <> 0 Then
      MsgBox "Erro C_7.1 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
      Valida_Chaves_Duplic = bRetorno
      Exit Function
    End If
    '
    rsDuplic.Close
    Set rsDuplic = Nothing
    '
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  '
  Set objGeral = Nothing
  bRetorno = True
  Valida_Chaves_Duplic = bRetorno
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Valida_Chaves_Duplic]", _
            Err.Description
End Function

Public Function Valida_Data_SO(ByVal pDataSO As String, _
                               ByVal strNivel As String) As Boolean
  On Error GoTo trata
  Dim sQtd_Vez    As String
  Dim dDtAtual    As Date
  Dim dDtPassada  As Date
  Dim sDtSO       As String
  Dim dDtSO       As Date
  Dim bRetorno    As Boolean
  '
  bRetorno = False
  'Data atual do sistema operacional + 10 minutos
  'dDtAtual = DateAdd("n", 10, Now)
  dDtAtual = DateAdd("n", 70, Now)
  'Data passada do sistema operacional - 70 minutos
  dDtPassada = DateAdd("n", -70, Now)
  'Última data gravada no banco de dados
  sDtSO = Encripta(pDataSO)
  sDtSO = Mid(sDtSO, 1, 2) & "/" & Mid(sDtSO, 4, 2) & "/" & Mid(sDtSO, 7, 4) & " " & Mid(sDtSO, 15, 2) & ":" & Mid(sDtSO, 17, 2) & ":" & Mid(sDtSO, 19, 2)
  dDtSO = CDate(Mid(sDtSO, 7, 4) & "/" & Mid(sDtSO, 4, 2) & "/" & Mid(sDtSO, 1, 2) & " " & Mid(sDtSO, 12, 8))
  '
  If dDtAtual < dDtSO Then
    If strNivel <> "ADM" Then
      MsgBox "Erro C_8.1 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
      Valida_Data_SO = bRetorno
      Exit Function
    End If
  End If
  '
  '
  If dDtPassada > dDtSO Then
    'If strNivel = "DIR" Or strNivel = "ADM" Or strNivel = "GER" Then
    '  If MsgBox("O último acesso ao sistema foi " & sDtSO & _
    '    vbCrLf & vbCrLf & "A data e hora do sistema operacional do windows foi alterada para " & Format(Now, "DD/MM/YYYY hh:mm") & " ou o sistema ficou paralizado neste período." & _
    '    vbCrLf & vbCrLf & "Tem certeza que a data e hora " & Format(Now, "DD/MM/YYYY hh:mm") & " é realmente a data e hora atual ? " & _
    '    vbCrLf & vbCrLf & "Caso seja selecionada a opção SIM, a data e hora " & Format(Now, "DD/MM/YYYY hh:mm") & " será gravada como data e hora do último acesso ao sistema." & _
    '    vbCrLf & vbCrLf & "Caso seja selecionada a opção NÃO, o sistema não prosseguirá. Modifique a data e hora do sistema operacional no windows e depois retorne ao sistema." & _
    '    vbCrLf & vbCrLf & "ATENÇÃO: Caso seja selecionada a opção SIM, e está não seja a data e hora atual, por proteção o sistema não permitirá que está data seja retrocedida. Por isso só selecione a opção SIM caso tenha certeza de que esa seja a data e hora atual." _
    '    , vbYesNo, gsTituloDll) = vbYes Then
    '
    '    Valida_Data_SO = True
    '    Exit Function
    '  Else
    '
    '    'MsgBox "Erro C_8.2 - Data do sistema operacional alterada fora do limite permitodo." & vbCrLf, vbOKOnly, gsTituloDll
    '    Valida_Data_SO = bRetorno
    '    Exit Function
    '  End If
    'Else 'É um operador
    '  MsgBox "Erro C_8.2 - Data do sistema operacional alterada fora do limite permitido." & vbCrLf & "Só o gerente conseguirá se logar no sistema !", vbOKOnly, gsTituloDll
    '  Valida_Data_SO = bRetorno
    '  Exit Function
    'End If
  End If
  '
'''    If sNivel = "DIR" Or sNivel = "ADM" Then
'''      If MsgBox("Erro D_1 - Acabou o prazo de validade de seu sistema." & vbCrLf & vbCrLf & vbCrLf & "Deseja entrar com a chave agora ?", vbYesNo, gsTituloDll) = vbYes Then
'''        bRetorno = False
'''        Valida_Prazo = bRetorno
'''        Exit Function
'''      Else
'''        MsgBox "Erro D_1 - Acabou o prazo de validade de seu sistema." & vbCrLf & "Contacte o suporte para adquirir uma nova chave !", vbOKOnly, gsTituloDll
'''        Valida_Prazo = bRetorno
'''      End If
'''    Else 'Não é autorizado a fornecer senha, mostra msg e sai, mas continua no sistema
'''      MsgBox "Erro D_1 - Acabou o prazo de validade de seu sistema." & vbCrLf & "Contacte o suporte para adquirir uma nova chave !", vbOKOnly, gsTituloDll
'''      bRetorno = True
'''      Valida_Prazo = bRetorno
'''    End If
  
  '
  bRetorno = True
  Valida_Data_SO = bRetorno
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Valida_Data_SO]", _
            Err.Description
End Function

Public Function Valida_Prazo(ByVal pDsn As String, _
                             ByVal pDataIni As String, _
                             ByVal sMostrarMsgExpira As String, _
                             ByVal sNivel As String, _
                             sEstourouPrazo As String) As Boolean
  On Error GoTo trata
  Dim sQtd_Vez        As String
  Dim dDtAtual        As Date
  Dim sDtIni          As String
  Dim dDtIni          As Date
  Dim dDataCompara    As Date
  Dim bRetorno        As Boolean
  Dim rs              As ADODB.Recordset
  Dim sSql            As String
  Dim lqtdMeses       As Long
  Dim lDifAtualpInicial   As Long
  Dim objGeral            As busSisMed.clsGeral
  '
  Set objGeral = New busSisMed.clsGeral
  '
  bRetorno = False
  '
  dDtAtual = Now
  sDtIni = Encripta(pDataIni)
  sDtIni = Mid(sDtIni, 1, 2) & "/" & Mid(sDtIni, 4, 2) & "/" & Mid(sDtIni, 7, 4)
  dDtIni = CDate(Mid(sDtIni, 7, 4) & "/" & Mid(sDtIni, 4, 2) & "/" & Mid(sDtIni, 1, 2) & " " & Mid(sDtIni, 12, 8))
  'Verifica Quantidade de meses a adicionar
  sSql = "Select Count(*) From Chave;"
  Set rs = objGeral.ExecutarSQL(sSql)
  lqtdMeses = rs(0)
  rs.Close
  Set rs = Nothing
  Set objGeral = Nothing
  'Adicionar a Data Inicial a quantidade de meses
  dDataCompara = DateAdd("m", lqtdMeses, dDtIni)
  'Adicionar a Data de comparação + 10 Dias
  dDataCompara = DateAdd("d", 10, dDataCompara)
  
  '
  If dDtAtual > dDataCompara Then
'    MsgBox "Erro D_1 - Acabou o prazo de validade de seu sistema." & vbCrLf & "Contacte o suporte para adquirir uma nova chave !", vbOKOnly, gsTituloDll
'    Valida_Prazo = bRetorno
'    Exit Function
    sEstourouPrazo = "S"
    If sNivel = "DIR" Or sNivel = "ADM" Then
      If MsgBox("Erro D_1 - Acabou o prazo de validade de seu sistema." & vbCrLf & vbCrLf & vbCrLf & "Deseja entrar com a chave agora ?", vbYesNo, gsTituloDll) = vbYes Then
        bRetorno = False
        Valida_Prazo = bRetorno
        Exit Function
      Else
        MsgBox "Erro D_1 - Acabou o prazo de validade de seu sistema." & vbCrLf & "Contacte o suporte para adquirir uma nova chave !", vbOKOnly, gsTituloDll
        Valida_Prazo = bRetorno
      End If
    Else 'Não é autorizado a fornecer senha, mostra msg e sai, mas continua no sistema
      MsgBox "Erro D_1 - Acabou o prazo de validade de seu sistema." & vbCrLf & "Contacte o suporte para adquirir uma nova chave !", vbOKOnly, gsTituloDll
      bRetorno = True
      Valida_Prazo = bRetorno
    End If


  Else
    sEstourouPrazo = "N"
    'Está no prazo, mas verifica se está na faixa dos dez dias
    lDifAtualpInicial = DateDiff("d", dDtAtual, dDataCompara)
    If lDifAtualpInicial < 10 Then
      If sMostrarMsgExpira = "S" Then
        'gsNivel = "DIR" Or Or gsNivel = "ADM")
        'Só mostra msg caso seja diretor ou Administrador
        If sNivel = "DIR" Or sNivel = "ADM" Then
          If MsgBox("Erro D_2 - Faltam " & lDifAtualpInicial & " dias para acabar o prazo de validade de seu sistema." & vbCrLf & vbCrLf & vbCrLf & "Deseja entrar com a chave agora ?", vbYesNo, gsTituloDll) = vbYes Then
            bRetorno = False
            Valida_Prazo = bRetorno
            Exit Function
          Else
            bRetorno = True
            Valida_Prazo = bRetorno
          End If
        Else 'Não é autorizado a fornecer senha, mostra msg e sai, mas continua no sistema
          MsgBox "Erro D_2 - Faltam " & lDifAtualpInicial & " dias para acabar o prazo de validade de seu sistema." & vbCrLf & "Contacte o Administrador do sistema para que ele forneça uma nova senha ! ", vbExclamation, gsTituloDll
          bRetorno = True
          Valida_Prazo = bRetorno
        End If
      Else
        bRetorno = True
        Valida_Prazo = bRetorno
      End If
      Exit Function
    Else
      'Ok
      bRetorno = True
    End If
  End If
  '
  bRetorno = True
  Valida_Prazo = bRetorno
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Valida_Prazo]", _
            Err.Description
End Function

Public Function Valida_Qtd(sQtd As String) As Boolean
  On Error GoTo trata
  Dim sQtd_Vez  As String
  Dim bRetorno  As Boolean
  '
  bRetorno = False
  '
  sQtd_Vez = Encripta(sQtd)
  If Len(Trim(sQtd_Vez)) <> 7 Then
    MsgBox "Erro C_4.1 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
    Valida_Qtd = bRetorno
    Exit Function
  End If
  '
  If Not IsNumeric(sQtd_Vez) Then
    MsgBox "Erro C_4.2 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
    Valida_Qtd = bRetorno
    Exit Function
  End If
  
  If Mid(sQtd_Vez, 7, 1) <> CalcDv(Mid(sQtd_Vez, 1, 6)) Then
    MsgBox "Erro C_4.3 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
    Valida_Qtd = bRetorno
    Exit Function
  End If
  bRetorno = True
  '
  Valida_Qtd = bRetorno
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Valida_Qtd]", _
            Err.Description
End Function

Public Sub Trata_Arquivo_Fisico(pApp As String)
  On Error GoTo trata
  Dim sAppPath    As String
  '
  sAppPath = pApp & "\Sismed.tmp"
  sAppPath = Dir(sAppPath)
  If Len(Trim(sAppPath)) <> 0 Then
    'Mata o arquivo fisicamene
    Kill pApp & "\Sismed.tmp"
  End If
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Trata_Arquivo_Fisico]", _
            Err.Description
End Sub
Public Sub Atualiza_Chave_Data_Atual(sDSN As String)
  On Error GoTo trata
  Dim sDataSO     As String
  Dim sHoraAtual  As String
  Dim sSql        As String
  Dim objGeral    As busSisMed.clsGeral
  '
  Set objGeral = New busSisMed.clsGeral
  '
  'Pega chaves do banco da tabela Gerencial
  '
  sDataSO = Format(Now, "DD") & CalcDv(Format(Now, "DD"))
  sDataSO = sDataSO & Format(Now, "MM") & CalcDv(Format(Now, "MM"))
  sDataSO = sDataSO & Format(Now, "YYYY") & CalcDv(Format(Now, "YYYY"))
  sDataSO = sDataSO & CalcDv(sDataSO)
  sDataSO = sDataSO & CalcDv(Mid(sDataSO, 1, 6))
  sDataSO = sDataSO & CalcDv(sDataSO)
  sHoraAtual = Format(Now, "hhmmss")
  sDataSO = sDataSO & sHoraAtual & CalcDv(sHoraAtual)
  '
  sSql = "UPDATE Gerencial set Chave3 = " & Formata_Dados(Encripta(sDataSO), tpDados_Texto, tpNulo_Aceita) & ";"
  objGeral.ExecutarSQLAtualizacao sSql
  '-------------------
  Set objGeral = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Atualiza_Chave_Data_Atual]", _
            Err.Description
End Sub
Public Sub Gravar_Chave(sDSN As String)
  On Error GoTo trata
  '
  GetVolumeInformation sRootPathName, sVolumeNameBuffer, lVolumeNameSize, _
    lVolSerNum, lMaximunComponentLenght, lFfileSystemFlags, sFfileNameSystemBuffer, lFileSystemNameSize
  lVolumeSerialNumber = Retorna_Volume_Info(lVolSerNum)
  '---------- Ok
  frmRegistrarChave.sEstourouPrazo = "S"
  If Not Adquirir_chave(lVolumeSerialNumber, sDSN) Then
    'Não validou Chave
    MsgBox "Erro C_1 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
  Else
    MsgBox "Chave registrada com sucesso"
  End If
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Gravar_Chave]", _
            Err.Description
End Sub

Public Function Valida_Chave(sDSN As String, _
                             sMostrarMsgExpira As String, _
                             sNivel As String) As Boolean
  On Error GoTo trata
  Dim bRetorno          As Boolean
  Dim rs                As ADODB.Recordset
  Dim rs1               As ADODB.Recordset
  Dim sSql              As String
  Dim sEstourouPrazo    As String
  Dim QtdChaves         As Long
  Dim objGeral          As busSisMed.clsGeral
  '
  Set objGeral = New busSisMed.clsGeral
Revalida_Tudo:
  bRetorno = False
  '
  'Pega chaves do banco da tabela Gerencial
  '
  sSql = "Select Chave2,Chave3, Chave4  from Gerencial;"
  Set rs = objGeral.ExecutarSQL(sSql)
  '
  If Not rs.EOF Then
    'Ok
    sSql = "Select Count(Chave1) from Chave;"
    Set rs1 = objGeral.ExecutarSQL(sSql)
    QtdChaves = rs1(0)
    rs1.Close
    Set rs1 = Nothing
    '
    GetVolumeInformation sRootPathName, sVolumeNameBuffer, lVolumeNameSize, _
      lVolSerNum, lMaximunComponentLenght, lFfileSystemFlags, sFfileNameSystemBuffer, lFileSystemNameSize
    lVolumeSerialNumber = Retorna_Volume_Info(lVolSerNum)
    '---------- Ok

    If Len(Trim(rs!Chave2 & "")) = 0 And Len(Trim(rs!Chave3 & "")) = 0 And Len(Trim(rs!Chave4 & "")) = 0 And QtdChaves = 0 Then
      'Tela para Pegar Chave pela primeira vez
      '-- Get VolNumber from Station
Adquirir_chave:
      
'''      Load frmRegistrarChave
'''      frmRegistrarChave.sVolumeKey = lVolumeSerialNumber
'''      frmRegistrarChave.sDataAtual = Format(Now, "DD/MM/YYYY")
'''      frmRegistrarChave.sBanco = sDSN
'''      frmRegistrarChave.Show vbModal
      
      If Not Adquirir_chave(lVolumeSerialNumber, sDSN) Then
        'Não validou Chave
        MsgBox "Erro C_1 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
      Else
        GoTo Revalida_Tudo
      End If
    Else
      'Existe chave, faz a verificação da tabela gerencial
      'Data Inicial
      '-----------
      'Validação se data da chave é data mesmo
      'e se data da chave bate com dvs
      If Not Valida_Data(rs!Chave2) Then
        MsgBox "Erro C_2 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
        Valida_Chave = bRetorno
        rs.Close
        Set rs = Nothing
        Set objGeral = Nothing
        Exit Function
      End If
      'Data Atual (última data gravada no banco de dados)
      '-----------
      'Validação se data atual da chave é data mesmo
      'e se data atual da chave bate com dvs
      If Not Valida_Data_Atual(rs!Chave3) Then
        MsgBox "Erro C_3 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
        Valida_Chave = bRetorno
        rs.Close
        Set rs = Nothing
        Set objGeral = Nothing
        Exit Function
      End If
      'Qtd vezes
      '-----------
      If Not Valida_Qtd(rs!Chave4) Then
        MsgBox "Erro C_4 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
        Valida_Chave = bRetorno
        rs.Close
        Set rs = Nothing
        Set objGeral = Nothing
        Exit Function
      End If
      'Qtd vezes = Nro Reg Chaves
      '-----------
      If CLng(QtdChaves) <> CLng(Encripta(Mid(rs!Chave4, 1, 6))) Then
        MsgBox "Erro C_5 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
        Valida_Chave = bRetorno
        rs.Close
        Set rs = Nothing
        Set objGeral = Nothing
        Exit Function
      End If
      'Continua validação, faz a verificação da tabela Chave
      '-----------------------------------------------------
      'Chaves validas
      '-----------
'      If Not Valida_Chaves_ok(lVolumeSerialNumber & "") Then
'        MsgBox "Erro C_6 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
'        Valida_Chave = bRetorno
'        Exit Function
'      End If
      'Duplicidades Chaves
      '-----------
      If Not Valida_Chaves_Duplic(sDSN) Then
        MsgBox "Erro C_7 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
        Valida_Chave = bRetorno
        rs.Close
        Set rs = Nothing
        Set objGeral = Nothing
        Exit Function
      End If
      '
      'Validar se data do SO menor ou Igual
      'a última data gravada no BD
      '-----------
      If Not Valida_Data_SO(rs!Chave3, sNivel) Then
        MsgBox "Erro C_8 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
        Valida_Chave = bRetorno
        rs.Close
        Set rs = Nothing
        Set objGeral = Nothing
        Exit Function
      End If
      'Validar Prazo do sistema
      If Not Valida_Prazo(sDSN, _
                          rs!Chave2, _
                          sMostrarMsgExpira, _
                          sNivel, _
                          sEstourouPrazo) Then
        frmRegistrarChave.sEstourouPrazo = "S"
        GoTo Adquirir_chave
      ElseIf sEstourouPrazo = "S" Then
        'MsgBox "Erro D_1 - Acabou o prazo de validade de seu sistema." & vbCrLf & "Contacte o suporte para adquirir uma nova chave !", vbOKOnly, gsTituloDll
        Valida_Chave = bRetorno
        rs.Close
        Set rs = Nothing
        Set objGeral = Nothing
        Exit Function
      End If
      'Ok
      bRetorno = True
    End If
  Else
    rs.Close
    Set rs = Nothing
    '-------------------
    MsgBox "Erro C_9 - A chave não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
    Valida_Chave = bRetorno
    Exit Function
  End If
  rs.Close
  Set rs = Nothing
  Set objGeral = Nothing
  '-------------------
  '
  Valida_Chave = bRetorno
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Valida_Chave]", _
            Err.Description
End Function

Public Function Valida_Primeira_Vez(sDSN As String, _
                                    ByVal pApp As String) As Boolean
  On Error GoTo trata
  Dim sSql              As String
  Dim sChave            As String
  Dim sChaveGerencial   As String
  Dim sChaveRegister    As String
  Dim rs                As ADODB.Recordset
  Dim bRetorno          As Boolean
  Dim sAppPath          As String
  Dim objGeral          As busSisMed.clsGeral
  '
  Set objGeral = New busSisMed.clsGeral
  '
  bRetorno = False
  'Pega chave do banco
  '
  sSql = "Select Count(*) from Chave;"
  Set rs = objGeral.ExecutarSQL(sSql)

  '
  If rs(0) = 0 Then
    sChave = "N"
  Else
    sChave = "S"
  End If
  rs.Close
  Set rs = Nothing
  '
  sSql = "Select Count(*) from Gerencial where Len(rtrim(chave2)) <> 0 ;"
  Set rs = objGeral.ExecutarSQL(sSql)
  '
  If rs(0) = 0 Then
    sChaveGerencial = "N"
  Else
    sChaveGerencial = "S"
  End If
  rs.Close
  Set rs = Nothing
  Set objGeral = Nothing
  '
  If sChaveGerencial = "N" And sChave = "N" Then
    'A chave estão em branco, verifica se register tb em branco
    'Get Register Information
    'sChaveRegister = Encripta(GetSetting(appname:=sAppName, section:=sSection, _
               Key:=sKey1))
    'If sChaveRegister = "" Then
    '  bRetorno = True
    'End If
    bRetorno = True
  End If
  '
  If bRetorno Then
    'Todas as chaves estão em branco, Verifica se existe
    'arquivo físico em windows system
    
    sAppPath = pApp & "\Sismed.tmp"
    sAppPath = Dir(sAppPath)
    If Len(Trim(sAppPath)) = 0 Then
      'Não encontrou o arquivo no diretório especificado
      bRetorno = False
      MsgBox "Erro A_1 - Chave inválida. Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
    Else
      'Encontrou o arquivo no diretório especificado
      'Mata o arquivo fisicamene
      'Kill ("C:\WINDOWS\SYSTEM\SISCLIE.DLL")
      bRetorno = True
    End If
  Else
    bRetorno = True
  End If
  '
  Valida_Primeira_Vez = bRetorno
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Valida_Primeira_Vez]", _
            Err.Description
End Function

Public Function Valida_Estacao(sDSN As String) As Boolean
  On Error GoTo trata
  Dim sSql        As String
  Dim rs1         As ADODB.Recordset
  Dim sChaveBanco As String
  '
  Dim bRetorno    As Boolean
  Dim sChave1     As String
  Dim objGeral    As busSisMed.clsGeral
  '
  Set objGeral = New busSisMed.clsGeral
  '
  bRetorno = False
  '-- Get VvolNumber from Station
  GetVolumeInformation sRootPathName, sVolumeNameBuffer, lVolumeNameSize, _
    lVolSerNum, lMaximunComponentLenght, lFfileSystemFlags, sFfileNameSystemBuffer, lFileSystemNameSize
  lVolumeSerialNumber = Retorna_Volume_Info(lVolSerNum)
  '---------- Ok
  'Get Register Information
  sChave1 = Encripta(GetSetting(appname:=sAppName, section:=sSection, _
             Key:=sKey1))
  'Pega chave do banco
  '
  sSql = "Select Chave1 from Gerencial;"
  Set rs1 = objGeral.ExecutarSQL(sSql)
  '
  If rs1.EOF Then
    sChaveBanco = ""
  Else
    sChaveBanco = rs1(0) & ""
  End If
  rs1.Close
  Set rs1 = Nothing
  Set objGeral = Nothing
  '-------------------
  'Validation Register
  If Len(Trim(sChave1)) = 0 Then 'Chave Invalida
    'Tela pra dar Get no Volume Key
    Load frmRegistrarEstacao
    frmRegistrarEstacao.sVolumeKey = lVolumeSerialNumber
    frmRegistrarEstacao.sBanco = sDSN
    frmRegistrarEstacao.Show vbModal
    If Not frmRegistrarEstacao.bRegistroValido Then
      'Não validou Chave
      MsgBox "Erro B_1 - A estação de trabalho não foi validada." & vbCrLf & "Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
    Else
      bRetorno = True
    End If
    Unload frmRegistrarEstacao
  ElseIf Encripta(sChave1) <> Encripta(Gerar_Chave_HD(lVolumeSerialNumber & "")) Then
    'Tentativa de Burlar o sistema
    MsgBox "Erro B_2 - Chave inválida. Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
  ElseIf InStr(1, sChaveBanco, Encripta(sChave1)) = 0 Then
    'Tentativa de Burlar o sistema
    MsgBox "Erro B_3 - Chave inválida. Contacte o suporte para maiores informações !", vbOKOnly, gsTituloDll
  Else
    'Get Record on DataBase
    'Validation Register Versus DataBase
    bRetorno = True
  End If
  Valida_Estacao = bRetorno
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[" & strClassName & ".Valida_Estacao]", _
            Err.Description
End Function


