Attribute VB_Name = "mdlUserFunction"
Option Explicit

Sub INCLUIR_VALOR_NO_CHECK(objCheck As CheckBox, _
                           blnValor As Variant)
  On Error GoTo trata
  If blnValor Then
    objCheck.Value = 1
  Else
    objCheck.Value = 0
  End If
  
  Exit Sub
trata:
  AmpN
  Err.Raise Err.Number, _
            "[mdlUserFunction.INCLUIR_VALOR_NO_CHECK]", _
            Err.Description
End Sub

'Proposito: Retornar último dia do mês
'Retorna DD
Public Function Retorna_ultimo_dia_do_mes(intMes As Integer, _
                                          intAno As Integer) As String
  On Error GoTo trata
  Dim intDia    As Integer
  Dim dtaTeste  As Date
  For intDia = 29 To 32
    dtaTeste = CDate(intAno & "/" & intMes & "/" & intDia)
  Next
  Exit Function
trata:
  Retorna_ultimo_dia_do_mes = Format(intDia - 1, "00")
End Function

'Propósito: Retornar e Gravar o Sequencial geral
Public Function RetornaGravaSequencial(Optional pCampo As String = "Sequencial") As Long
  On Error GoTo trata
  '
  Dim objRs       As ADODB.Recordset
  Dim strSql      As String
  Dim lngSeq      As Long
  Dim objGeral    As busSisMed.clsGeral
  '
  Set objGeral = New busSisMed.clsGeral
  '
  strSql = "SELECT ISNULL(MAX(" & pCampo & "),0) + 1 AS SEQ " & _
    "From SEQUENCIAL"
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  If objRs.EOF Then
    lngSeq = 1
  ElseIf Not IsNumeric(objRs.Fields("SEQ").Value) Then
    lngSeq = 1
  Else
    lngSeq = objRs.Fields("SEQ").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'Caso tenha atingido o limite do sequencial, voltar a 1
  If lngSeq > glMaxSequencialImp Then
    lngSeq = 1
  End If
  
  strSql = "Select Count(*) As TOT from Sequencial"
  Set objRs = objGeral.ExecutarSQL(strSql)
  If objRs.Fields("TOT").Value = 0 Then
    strSql = "INSERT INTO SEQUENCIAL (" & pCampo & ") VALUES (" & lngSeq & ")"
  Else
    strSql = "UPDATE SEQUENCIAL SET " & pCampo & " = " & lngSeq
  End If
  
  '
  objRs.Close
  Set objRs = Nothing
  '---
  objGeral.ExecutarSQLAtualizacao strSql
  '
  RetornaGravaSequencial = CLng(lngSeq)
  '
  Set objGeral = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.RetornaGravaSequencial]", _
            Err.Description
End Function


Public Sub TratarStatus(strStatus As String, _
                        strStatusImp As String, _
                        ByRef lblCor1 As Label, _
                        ByRef lblCor2 As Label)
  On Error GoTo trata
  '
  Select Case strStatus
  Case "I"
    'Amarelo
    lblCor1.BackColor = &H80FFFF
    lblCor1.Caption = "Inicial"
  Case "M"
    'Azul
    lblCor1.BackColor = &HFFFF80
    lblCor1.Caption = "Movimento após o fechamento"
  Case "F"
    'Verde
    lblCor1.BackColor = &H80FF80
    lblCor1.Caption = "Fechada"
  
  Case "C"
    'Vermelho
    lblCor1.BackColor = &HFF&
    lblCor1.Caption = "Cancelada"
  Case Else
    lblCor1.BackColor = &H8000000F
    lblCor1.Caption = "Indefinido"
  End Select
  '
  Select Case strStatusImp
  Case "S"
    'Verde
    lblCor2.BackColor = &H80FF80
    lblCor2.Caption = "Sim"
  Case "N"
    'Vermelho
    lblCor2.BackColor = &HFF&
    lblCor2.Caption = "Não"
  Case Else
    lblCor2.BackColor = &H8000000F
    lblCor2.Caption = "Ind"
  End Select
  
  
  Exit Sub
trata:
  TratarErro Err.Number, _
             "[mdlUserFunction.Main]", _
             Err.Description
End Sub


Public Sub Main()
  On Error GoTo trata
  '
  frmUserSplash.Show
  frmUserSplash.Refresh
  '
  frmUserLogin.QuemChamou = 0
  Load frmUserLogin
  frmUserLogin.Show vbModal
  '
  Exit Sub
trata:
  TratarErro Err.Number, _
             "[mdlUserFunction.Main]", _
             Err.Description
  End
End Sub
Public Function Valida_Hora(pMsk As MaskEdBox, _
                            pTipo As TpObriga, _
                            Optional blnSetarFocoControle As Boolean = False) As Boolean
  On Error GoTo trata
  Dim EData As Boolean
  EData = True
  'Verifica se Mmaskedit é data
  If Not IsDate("01/01/1900 " & pMsk.Text) Then
    EData = False
  Else
    'If CInt(Mid(pMsk.ClipText, 3, 2)) > 12 Then
    '  EData = False
    'End If
  End If
  If pTipo = TpObrigatorio And Not (EData) Then
    'Campo é obrigatório e não é data
    Valida_Hora = False
  ElseIf pTipo = TpNaoObrigatorio Then
    'Campo não é obrigatório
    If Len(pMsk.ClipText) <> 0 And Not EData Then
      'Digitou algo que não é data
      Valida_Hora = False
    Else
      Valida_Hora = True
    End If
  Else
    Valida_Hora = True
  End If
  If Valida_Hora = False Then
    Pintar_Controle pMsk, tpCorContr_Erro
    If blnSetarFocoControle Then
      SetarFoco pMsk
      blnSetarFocoControle = False
    End If
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.Valida_Hora]", _
            Err.Description
End Function

Public Sub TratarErro(ByVal pNumero As Long, _
                    ByVal pDescricao As String, _
                    ByVal pSource As String)
  '
  On Error Resume Next
  Dim strUsuario As String
  Dim intI       As Integer
    
  intI = FreeFile
  Open App.Path & "\SisMed.txt" For Append As #intI
  
  Print #intI, Format(Now(), "DD/MM/YYYY hh:mm") & ";" & pNumero & ";" & pSource & ";" & pDescricao
  Close #intI
  'mostrar Mensagem
  MsgBox "O Seginte Erro Ocorreu: " & vbCrLf & vbCrLf & _
    "Número: " & pNumero & vbCrLf & _
    "Descrição: " & pDescricao & vbCrLf & vbCrLf & _
    "Origem: " & pSource & vbCrLf & _
    "Data/Hora: " & Format(Now(), "DD/MM/YYYY hh:mm") & vbCrLf & _
    "Erro gravado no arquivo: " & App.Path & "\SisMed.txt" & vbCrLf & vbCrLf & _
    "Caso o erro persista contacte o suporte e envie o arquivo acima, informando a data e hora acima informada da ocorrência deste erro.", vbCritical, TITULOSISTEMA
End Sub



'Propósito Abrir Registros do Sistema
'recebe parametro pAcao que assume
'0 - Captura parametros inicias
'1 -  grava últ usuário que
'2 - Grava BMP
'acessou o sistema
'Caso algum parametro seja nulo, regrava no Registro
Public Function CapturaParametrosRegistro(pAcao As Integer)
  Dim iLenCaminho, iLenArquivo As Long
  Dim iRet
  On Error GoTo RotErro
Repeticao:
  AmpS
  Select Case pAcao
  Case 0
    'Captura Banco de Dados
    gsBDadosPath = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:="CaminhoDB")
    If Len(Trim(gsBDadosPath)) = 0 Then
      'Registro não esta gravado está em branco
      iRet = MsgBox("Não foi possível localizar o banco de dados " & _
        nomeBDados & ". Você deseja localizá-lo manualmente?", _
        vbQuestion + vbYesNo, TITULOSISTEMA)
      If iRet = vbYes Then
        frmMDI.CommonDialog1.ShowOpen
        iLenCaminho = Len(frmMDI.CommonDialog1.FileName)
        iLenArquivo = Len(frmMDI.CommonDialog1.FileTitle)
        gsBDadosPath = Left(frmMDI.CommonDialog1.FileName, iLenCaminho - iLenArquivo)
        SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                  Key:="CaminhoDB", setting:=gsBDadosPath
        GoTo Repeticao 'Captura Novamente os parametros até encontrar
                        'Parametros válidos
      Else
        End
      End If
    End If
    'Captura Caminho Crystal
    gsReportPath = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:="CaminhoReport")
    If Len(Trim(gsReportPath)) = 0 Then
      'Registro não esta gravado está em branco
      iRet = MsgBox("Não foi possível localizar o caminho dos Formulários (.RPT)." & _
        " Você deseja localizá-lo manualmente?", _
        vbQuestion + vbYesNo, TITULOSISTEMA)
      If iRet = vbYes Then
        frmMDI.CommonDialog1.ShowOpen
        iLenCaminho = Len(frmMDI.CommonDialog1.FileName)
        iLenArquivo = Len(frmMDI.CommonDialog1.FileTitle)
        gsReportPath = Left(frmMDI.CommonDialog1.FileName, iLenCaminho - iLenArquivo)
        SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                  Key:="CaminhoReport", setting:=gsReportPath
        GoTo Repeticao 'Captura Novamente os parametros até encontrar
                        'Parametros válidos
      Else
        End
      End If
    End If
    'Captura Caminho App
    gsAppPath = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:="CaminhoApp")
    If Len(Trim(gsAppPath)) = 0 Then
      'Registro não esta gravado está em branco
      iRet = MsgBox("Não foi possível localizar o caminho do Aplicativo (.EXE)." & _
        " Você deseja localizá-lo manualmente?", _
        vbQuestion + vbYesNo, TITULOSISTEMA)
      If iRet = vbYes Then
        frmMDI.CommonDialog1.ShowOpen
        iLenCaminho = Len(frmMDI.CommonDialog1.FileName)
        iLenArquivo = Len(frmMDI.CommonDialog1.FileTitle)
        gsAppPath = Left(frmMDI.CommonDialog1.FileName, iLenCaminho - iLenArquivo)
        SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                  Key:="CaminhoApp", setting:=gsAppPath
        GoTo Repeticao 'Captura Novamente os parametros até encontrar
                        'Parametros válidos
      Else
        End
      End If
    End If
    'Captura Nome do Usuário
    gsNomeUsu = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:="Usuario")
    'Captura o Nome do Curso
    gsNomeEmpresa = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:="Empresa")
    If Len(Trim(gsNomeEmpresa)) = 0 Then
      'Registro não esta gravado está em branco
      gsNomeEmpresa = "XXX"
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="Empresa", setting:=gsNomeEmpresa
      GoTo Repeticao 'Captura Novamente os parametros até encontrar
                      'Parametros válidos
    End If
    'Captura o Caminho dos bitmaps
    gsBMPPath = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:="CaminhoBMP")
    If Len(Trim(gsBMPPath)) = 0 Then
      'Registro não esta gravado está em branco
      gsBMPPath = gsBDadosPath & "Images\BMP\"
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="CaminhoBMP", setting:=gsBMPPath
      GoTo Repeticao 'Captura Novamente os parametros até encontrar
                      'Parametros válidos
    End If
    'Captura o Caminho dos Icons
    gsIconsPath = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:="CaminhoIcons")
    If Len(Trim(gsIconsPath)) = 0 Then
      'Registro não esta gravado está em branco
      gsIconsPath = gsBDadosPath & "Images\Icons\"
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="CaminhoIcons", setting:=gsIconsPath
      GoTo Repeticao 'Captura Novamente os parametros até encontrar
                      'Parametros válidos
    End If
    'Captura o Nome do BitMap
    gsBMP = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:="BMP")
    'Captura o caminho do BackUp
    gsPathBackup = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:="CaminhoBackUp")
    If Len(Trim(gsPathBackup)) = 0 Then
      'Registro não esta gravado está em branco
      gsPathBackup = gsAppPath & "BackUp\"
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="CaminhoBackUp", setting:=gsPathBackup
      GoTo Repeticao 'Captura Novamente os parametros até encontrar
                      'Parametros válidos
    End If
  Case 1
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="Usuario", setting:=gsNomeUsu
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="Nivel", setting:=gsNivel
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="Ativo", setting:="S"

  Case 2
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="BMP", setting:=gsBMP
  Case 3
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="Nivel", setting:=""
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="Ativo", setting:="N"
  
  
  End Select
  AmpN
  Exit Function
RotErro:
  AmpN
  frmMDI.CommonDialog1.ShowOpen
  iLenCaminho = Len(frmMDI.CommonDialog1.FileName)
  iLenArquivo = Len(frmMDI.CommonDialog1.FileTitle)
  gsBDadosPath = Left(frmMDI.CommonDialog1.FileName, iLenCaminho - iLenArquivo)
  SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
            Key:="CaminhoDB", setting:=gsBDadosPath
  
  AmpN
End Function


Public Sub AmpS()
  On Error Resume Next
  Screen.MousePointer = vbHourglass
End Sub

Public Sub AmpN()
  On Error Resume Next
  Screen.MousePointer = vbDefault
End Sub

'Propósito: Montar menu, de acordo com o nível de acesso
Public Sub Monta_Menu(pAcao As Integer)
  'pAcao  Assume 0 Desconexão 1 Conexão
  On Error GoTo trata
  '
  'Desabilita Menu
  frmMDI.snuDiretoria(0).Visible = False
  frmMDI.snuGerencia(0).Visible = False
  frmMDI.snuCaixa(0).Visible = False
  frmMDI.snuLaboratorio(0).Visible = False
  frmMDI.snuFinanceiro(0).Visible = False
  frmMDI.snuRelatorio(0).Visible = False
  frmMDI.snuArquivista(0).Visible = False
  frmMDI.snuAtendimento(0).Visible = False
  '
  frmMDI.mnuArquivo(2).Visible = False
  frmMDI.mnuArquivo(3).Visible = False
  frmMDI.mnuArquivo(4).Visible = False
  '
  If pAcao = 1 Then 'Monta Conexão
    Select Case Trim(gsNivel)
    Case gsAdmin
      'Acesso Completo
      frmMDI.snuDiretoria(0).Visible = True
      frmMDI.snuGerencia(0).Visible = True
      frmMDI.snuCaixa(0).Visible = True
      frmMDI.snuLaboratorio(0).Visible = True
      frmMDI.snuFinanceiro(0).Visible = True
      frmMDI.snuRelatorio(0).Visible = True
      frmMDI.snuRelFinanc(0).Visible = True
      frmMDI.snuRelFinanc(1).Visible = True
      frmMDI.snuArquivista(0).Visible = True
      frmMDI.snuAtendimento(0).Visible = True
      '
      frmMDI.mnuArquivo(2).Visible = True
      frmMDI.mnuArquivo(3).Visible = True
      frmMDI.mnuArquivo(4).Visible = True
    Case gsDiretor
      'Diretoria
      frmMDI.snuDiretoria(0).Visible = True
      frmMDI.snuGerencia(0).Visible = False
      frmMDI.snuCaixa(0).Visible = False
      frmMDI.snuLaboratorio(0).Visible = False
      frmMDI.snuFinanceiro(0).Visible = False
      frmMDI.snuRelatorio(0).Visible = True
      frmMDI.snuRelFinanc(0).Visible = True
      frmMDI.snuRelFinanc(1).Visible = True
      frmMDI.snuArquivista(0).Visible = True
      frmMDI.snuAtendimento(0).Visible = False
      '
      frmMDI.mnuArquivo(2).Visible = True
      frmMDI.mnuArquivo(3).Visible = True
      frmMDI.mnuArquivo(4).Visible = True
    Case gsGerente
      'Gerencia
      frmMDI.snuDiretoria(0).Visible = False
      frmMDI.snuGerencia(0).Visible = True
      frmMDI.snuCaixa(0).Visible = False
      frmMDI.snuLaboratorio(0).Visible = False
      frmMDI.snuFinanceiro(0).Visible = False
      frmMDI.snuRelatorio(0).Visible = True
      frmMDI.snuRelFinanc(0).Visible = True
      frmMDI.snuRelFinanc(1).Visible = False
      frmMDI.snuArquivista(0).Visible = False
      frmMDI.snuAtendimento(0).Visible = False
      '
      frmMDI.mnuArquivo(2).Visible = True
      frmMDI.mnuArquivo(3).Visible = True
      frmMDI.mnuArquivo(4).Visible = True
    Case gsCaixa
      'Caixa
      frmMDI.snuDiretoria(0).Visible = False
      frmMDI.snuGerencia(0).Visible = False
      frmMDI.snuCaixa(0).Visible = True
      frmMDI.snuLaboratorio(0).Visible = False
      frmMDI.snuFinanceiro(0).Visible = False
      frmMDI.snuRelatorio(0).Visible = False
      frmMDI.snuArquivista(0).Visible = False
      frmMDI.snuAtendimento(0).Visible = False
      '
      frmMDI.mnuArquivo(2).Visible = False
      frmMDI.mnuArquivo(3).Visible = True
      frmMDI.mnuArquivo(4).Visible = True
    Case gsLaboratorio
      'Laboratorio
      frmMDI.snuDiretoria(0).Visible = False
      frmMDI.snuGerencia(0).Visible = False
      frmMDI.snuCaixa(0).Visible = False
      frmMDI.snuLaboratorio(0).Visible = True
      frmMDI.snuFinanceiro(0).Visible = False
      frmMDI.snuRelatorio(0).Visible = False
      frmMDI.snuArquivista(0).Visible = False
      frmMDI.snuAtendimento(0).Visible = False
      '
      frmMDI.mnuArquivo(2).Visible = False
      frmMDI.mnuArquivo(3).Visible = True
      frmMDI.mnuArquivo(4).Visible = True
    Case gsFinanceiro
      'Financeiro
      frmMDI.snuDiretoria(0).Visible = False
      frmMDI.snuGerencia(0).Visible = False
      frmMDI.snuCaixa(0).Visible = False
      frmMDI.snuLaboratorio(0).Visible = False
      frmMDI.snuFinanceiro(0).Visible = True
      frmMDI.snuRelatorio(0).Visible = True
      frmMDI.snuRelFinanc(0).Visible = True
      frmMDI.snuRelFinanc(1).Visible = False
      frmMDI.snuArquivista(0).Visible = False
      frmMDI.snuAtendimento(0).Visible = False
      '
      frmMDI.mnuArquivo(2).Visible = False
      frmMDI.mnuArquivo(3).Visible = True
      frmMDI.mnuArquivo(4).Visible = True
    Case gsArquivista
      'Arquivista
      frmMDI.snuDiretoria(0).Visible = False
      frmMDI.snuGerencia(0).Visible = False
      frmMDI.snuCaixa(0).Visible = False
      frmMDI.snuLaboratorio(0).Visible = False
      frmMDI.snuFinanceiro(0).Visible = False
      frmMDI.snuRelatorio(0).Visible = False
      frmMDI.snuArquivista(0).Visible = True
      frmMDI.snuAtendimento(0).Visible = False
      '
      frmMDI.mnuArquivo(2).Visible = False
      frmMDI.mnuArquivo(3).Visible = True
      frmMDI.mnuArquivo(4).Visible = True
    
    Case gsPrestador
      'Prestador
      frmMDI.snuDiretoria(0).Visible = False
      frmMDI.snuGerencia(0).Visible = False
      frmMDI.snuCaixa(0).Visible = False
      frmMDI.snuLaboratorio(0).Visible = False
      frmMDI.snuFinanceiro(0).Visible = False
      frmMDI.snuRelatorio(0).Visible = False
      frmMDI.snuArquivista(0).Visible = False
      frmMDI.snuAtendimento(0).Visible = True
      '
      frmMDI.mnuArquivo(2).Visible = False
      frmMDI.mnuArquivo(3).Visible = True
      frmMDI.mnuArquivo(4).Visible = True


    End Select
  End If
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.Monta_Menu]", _
            Err.Description
End Sub

Public Function Valida_String(objControl As Object, _
                              TpTipo As TpObriga, _
                              Optional blnSetarFocoControle As Boolean = False) As Boolean
  On Error GoTo trata
  Dim blnValido As Boolean
  'Verifica se Textbox é Válido
  If Len(Trim(objControl.Text)) <> 0 Then
    blnValido = True
  End If
  If TpTipo = TpObrigatorio And Not (blnValido) Then
    'Campo é obrigatório e não é Valor
    Valida_String = False
  Else
    Valida_String = True
  End If
  If Not Valida_String Then
    Pintar_Controle objControl, tpCorContr_Erro
    If blnSetarFocoControle Then
      SetarFoco objControl
      blnSetarFocoControle = False
    End If
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.Valida_String]", _
            Err.Description
End Function


Public Sub Pintar_Controle(pControle As Variant, _
                            pCor As tpCorControle)
  On Error Resume Next
  AmpS
  pControle.BackColor = pCor
  AmpN
End Sub

Public Sub SetarFoco(objTarget As Object)
  On Error Resume Next
  If objTarget.Visible = True And objTarget.Enabled = True Then
    objTarget.SetFocus
  End If
End Sub


Public Sub TratarErroPrevisto(ByVal pDescricao As String, _
                              Optional pSource As String = "")
  '
  On Error Resume Next
  'mostrar Mensagem
  MsgBox "Erro(s): " & vbCrLf & vbCrLf & _
    pDescricao '& vbCrLf & vbCrLf '& _
    '"Módulo: " & pSource & vbCrLf & vbCrLf & _
    '"Reavalie as informações e corrija os dados para que a alteração seja efetivada.", vbExclamation, TITULOSISTEMA
End Sub


'Propósito: criptografar a senha do usuário armazenada no banco de dados
'Entrada: senha
'Retorna: senha
          'caso entrada seja não criptografada a saída é criptografada e vice-versa

Public Function Encripta(Senha As String) As String
  On Error GoTo trata
  Dim intI As Integer
  Dim strRetorno As String
  For intI = 1 To Len(Senha)
    strRetorno = Mid(Senha, intI, 1)
    strRetorno = 255 - Asc(strRetorno)
    Encripta = Encripta & Chr(strRetorno)
  Next intI
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.Encripta]", _
            Err.Description
End Function

'Propósito: Centralizar um form MDI Child no form MDI.
'Entradas:  frmCenter - Form a centralizar

Public Sub CenterForm(frmCenter As Form)
  Dim intHeight      As Integer
  Dim intLeftOffset  As Integer
  Dim intTop         As Integer
  Dim intWidth       As Integer
  Dim intLeft        As Integer
  Dim intTopOffset   As Integer
  '
  On Error GoTo trata
  AmpS
  If frmCenter.MDIChild = True Then
    intHeight = frmMDI.ScaleHeight
    intWidth = frmMDI.ScaleWidth
    intTop = frmMDI.Top
    intLeft = frmMDI.Left
  Else
    intHeight = Screen.Height
    intWidth = Screen.Width
    intTop = 0
    intLeft = 0
  End If

  'Calcula "left offset"
  intLeftOffset = ((intWidth - frmCenter.Width) / 2) + intLeft
  If (intLeftOffset + frmCenter.Width > Screen.Width) Or intLeftOffset < 100 Then
    intLeftOffset = 100
  End If

  'Calcula "top offset"
  intTopOffset = ((intHeight - frmCenter.Height) / 2) + intTop
  If (intTopOffset + frmCenter.Height > Screen.Height) Or intTopOffset < 100 Then
    intTopOffset = 100
  End If
  'Centraliza o form
  frmCenter.Move intLeftOffset, intTopOffset
  AmpN
  Exit Sub
trata:
  AmpN
  Err.Raise Err.Number, _
            "[mdlUserFunction.CenterForm]", _
            Err.Description
End Sub


Public Sub LerFiguras(pForm As Form, pOp As tpBmpForm, Optional pbtnOK As CommandButton, Optional pbtnCancelar As CommandButton, Optional pbtnFechar As CommandButton, Optional pbtnExcluir As CommandButton, Optional pbtnSenha As CommandButton, Optional pbtnIncluir As CommandButton, Optional pbtnAlterar As CommandButton, Optional pbtnFiltrar As CommandButton, Optional pbtnImprimir As CommandButton, Optional pbtnConsultar As CommandButton)
  On Error GoTo trata
    
  If pOp = tpBmpForm.tpBmp_Login Then
    pForm.Picture = LoadPicture(gsBMPPath & "Fundo.jpg")
  ElseIf pOp = tpBmpForm.tpBmp_MDI Then
    pForm.Picture = LoadPicture(gsBMPPath & "Fundo.jpg")
    pForm.Icon = LoadPicture(gsIconsPath & "Logo.ico")
  Else
    pForm.Icon = LoadPicture(gsIconsPath & "areatrab.ico")
  End If
  
  If Not (pbtnConsultar Is Nothing) Then
    pbtnConsultar.Picture = LoadPicture(gsIconsPath & "Procurar.ico")
    pbtnConsultar.DownPicture = LoadPicture(gsIconsPath & "ProcurarDown.ico")
    pbtnConsultar.ToolTipText = "Consultar"
  End If
  If Not (pbtnImprimir Is Nothing) Then
    pbtnImprimir.Picture = LoadPicture(gsIconsPath & "Impressora.ico")
    pbtnImprimir.DownPicture = LoadPicture(gsIconsPath & "ImpressoraDown.ico")
    pbtnImprimir.ToolTipText = "Imprimir"
  End If
  If Not (pbtnOK Is Nothing) Then
    pbtnOK.Picture = LoadPicture(gsIconsPath & "Ok.ico")
    pbtnOK.DownPicture = LoadPicture(gsIconsPath & "OkDown.ico")
    pbtnOK.ToolTipText = "Ok"
  End If
  If Not (pbtnAlterar Is Nothing) Then
    pbtnAlterar.Picture = LoadPicture(gsIconsPath & "Alterar.ico")
    pbtnAlterar.DownPicture = LoadPicture(gsIconsPath & "AlterarDown.ico")
    pbtnAlterar.ToolTipText = "Alterar"
  End If
  
  If Not (pbtnIncluir Is Nothing) Then
    pbtnIncluir.Picture = LoadPicture(gsIconsPath & "Incluir.ico")
    pbtnIncluir.DownPicture = LoadPicture(gsIconsPath & "IncluirDown.ico")
    pbtnIncluir.ToolTipText = "Incluir"
  End If
  
  If Not (pbtnCancelar Is Nothing) Then
    pbtnCancelar.Picture = LoadPicture(gsIconsPath & "Cancelar.ico")
    pbtnCancelar.DownPicture = LoadPicture(gsIconsPath & "CancelarDown.ico")
    pbtnCancelar.ToolTipText = "Cancelar"
  End If
  If Not (pbtnExcluir Is Nothing) Then
    pbtnExcluir.Picture = LoadPicture(gsIconsPath & "Excluir.ico")
    pbtnExcluir.DownPicture = LoadPicture(gsIconsPath & "ExcluirDown.ico")
    pbtnExcluir.ToolTipText = "Excluir"
  End If
  If Not (pbtnSenha Is Nothing) Then
    pbtnSenha.Picture = LoadPicture(gsIconsPath & "Senha.ico")
    pbtnSenha.DownPicture = LoadPicture(gsIconsPath & "SenhaDown.ico")
    pbtnSenha.ToolTipText = "Senha"
  End If
  If Not (pbtnFechar Is Nothing) Then
    pbtnFechar.Picture = LoadPicture(gsIconsPath & "Sair.ico")
    pbtnFechar.DownPicture = LoadPicture(gsIconsPath & "SairDown.ico")
    pbtnFechar.ToolTipText = "Sair"
  End If
  If Not (pbtnFiltrar Is Nothing) Then
    pbtnFiltrar.Picture = LoadPicture(gsIconsPath & "Filtrar.ico")
    pbtnFiltrar.DownPicture = LoadPicture(gsIconsPath & "FiltrarDown.ico")
    pbtnFiltrar.ToolTipText = "Aplicar Filtro"
  End If
  '
  Exit Sub
trata:
  AmpN
  Err.Raise Err.Number, _
            "[mdlUserFunction.LerFiguras]", _
            Err.Description
End Sub

Public Sub Seleciona_Conteudo_Controle(Controle As Object)
  On Error GoTo trata
  Controle.SelStart = 0
  Controle.SelLength = Len(Controle.Text)
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Public Function GetRelativeBookmarkGeral(Bookmark As Variant, _
        Offset As Long, intLINHASMATRIZ As Long) As Variant
  ' GetRelativeBookmark is used to get a bookmark for a
  ' row that is a specified number of rows away from the
  ' given row. Offset specifies the number of rows to
  ' move. A positive Offset indicates that the desired
  ' row is after the one referred to by Bookmark, and a
  ' negative Offset means it is before the one referred
  ' to by Bookmark.
  On Error GoTo trata
  Dim Index As Long
    
  ' Compute the row index for the desired row
  Index = IndexFromBookmarkGeral(Bookmark, Offset, intLINHASMATRIZ)
  If Index < 0 Or Index >= intLINHASMATRIZ Then
    ' Index refers to a row before the first or after
    ' the last, so just return Null.
    GetRelativeBookmarkGeral = Null
  Else
    GetRelativeBookmarkGeral = MakeBookmarkGeral(Index)
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.GetRelativeBookmarkGeral]", _
            Err.Description
End Function



Public Function IndexFromBookmarkGeral(Bookmark As Variant, _
        Offset As Long, intLINHASMATRIZ As Long) As Long
  ' This support function is used only by the remaining
  ' support functions. It is not used directly by the
  ' unbound events.
  ' IndexFromBookmark computes the row index that
  ' corresponds to a row that is (Offset) rows from the
  ' row specified by the Bookmark parameter. For example,
  ' if Bookmark refers to the index 50 of the dataset
  ' array and Offset = -10, then IndexFromBookmark will
  ' return 50 + (-10), or 40. Thus to get the index of
  ' the row specified by the bookmark itself, call
  ' IndexFromBookmark with an Offset of 0. If the given
  ' Bookmark is Null, it refers to BOF or EOF. If
  ' Offset < 0, the grid is requesting rows before the
  ' row specified by Bookmark, and so we must be at EOF
  ' because prior rows do not exist for BOF. Conversely,
  ' if Offset > 0, we are at BOF.
  On Error GoTo trata
  Dim Index As Long
  If IsNull(Bookmark) Then
    If Offset < 0 Then
      ' Bookmark refers to EOF. Since (MaxRow - 1)
      ' is the index of the last record, we can use
      ' an index of (MaxRow) to represent EOF.
      Index = intLINHASMATRIZ + Offset
    Else
      ' Bookmark refers to BOF. Since 0 is the index
      ' of the first record, we can use an index of
      ' -1 to represent BOF.
      Index = -1 + Offset
    End If
  Else
    ' Convert string to long integer
    Index = Val(Bookmark) + Offset
  End If
    
  ' Check to see if the row index is valid:
  '   (0 <= Index < MaxRow).
  ' If not, set it to a large negative number to
  ' indicate that it is invalid.
  If Index >= 0 And Index < intLINHASMATRIZ Then
    IndexFromBookmarkGeral = Index
  Else
    IndexFromBookmarkGeral = -9999
  End If
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.IndexFromBookmarkGeral]", _
            Err.Description
End Function

Public Function MakeBookmarkGeral(Index As Long) As Variant
  ' This support function is used only by the remaining
  ' support functions. It is not used directly by the
  ' unbound events. It is a good idea to create a
  ' MakeBookmark function such that all bookmarks can be
  ' created in the same way. Thus the method by which
  ' bookmarks are created is consistent and easy to
  ' modify. This function creates a bookmark when given
  ' an array row index.
  ' Since we have data stored in an array, we will just
  ' use the array index as our bookmark. We will convert
  ' it to a string first, using the CStr function.
  On Error GoTo trata
  MakeBookmarkGeral = CStr(Index)
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.MakeBookmarkGeral]", _
            Err.Description
End Function


Public Sub Selecionar_Conteudo(pControle As Variant)
  On Error Resume Next
  AmpS
  pControle.SelStart = 0
  pControle.SelLength = Len(pControle)
  AmpN
End Sub


Function INCLUIR_VALOR_NO_MASK(pMask As MaskEdBox, pValor As Variant, pTipo As tpMaskValor) As String
  On Error GoTo trata
  Dim strMask As String
  'Limpa Maskedit
  With pMask
    strMask = .Mask
    .Mask = ""
    .Text = ""
    .Mask = strMask
  End With
  '
  If pTipo = tpMaskValor.TpMaskData Then
    'Campo é data
    If Len(strMask) = 10 Then
      'Formato DD/MM/YYYY
      If Not IsNull(pValor) And pValor & "" <> "" Then pMask.Text = Format(pValor, "DD/MM/YYYY")
    ElseIf Len(strMask) = 16 Then
      'Formato DD/MM/YYYY hh:mm
      If Not IsNull(pValor) Then pMask.Text = Format(pValor, "DD/MM/YYYY hh:mm")
    ElseIf Len(strMask) = 5 Then
      'Formato hh:mm
      If Not IsNull(pValor) Then pMask.Text = Format(pValor, "hh:mm")
    ElseIf Len(strMask) = 8 Then
      'Formato hh:mm:ss
      If Not IsNull(pValor) Then pMask.Text = Format(pValor, "hh:mm:ss")
    End If
  ElseIf pTipo = tpMaskValor.TpMaskLongo Then
    'Campo é Longo
    If Not IsNull(pValor) Then pMask.Text = Format(pValor, "#,##0")
    
  ElseIf pTipo = tpMaskValor.TpMaskMoeda Then
    'Campo é moeda
    If Not IsNull(pValor) Then pMask.Text = Format(pValor, "#,##0.00##")
  ElseIf pTipo = tpMaskValor.TpMaskOutros Then
    'Campo é outros
    If Not IsNull(pValor) Then
      If Len(Trim(pValor)) <> 0 Then pMask.Text = pValor
    End If
  ElseIf pTipo = tpMaskValor.TpMaskSemMascara Then
    'Campo é guardado sem máscara
    If Not IsNull(pValor) Then
      If Len(Trim(pValor)) <> 0 Then
        With pMask
          strMask = .Mask
          .Mask = ""
          .Text = AplicarMascara(pValor, strMask)
          .Mask = strMask
        End With
      End If
    End If
  
  End If
  
  Exit Function
trata:
  AmpN
  Err.Raise Err.Number, _
            "[mdlUserFunction.INCLUIR_VALOR_NO_MASK]", _
            Err.Description
End Function


Public Function AplicarMascara(strTexto, strMascara) As String
  On Error GoTo trata
  Dim intQtdCaracterMasc  As Integer
  Dim intX                As Integer
  Dim strRetorno          As String
  strRetorno = ""
  intX = 0
  For intQtdCaracterMasc = 1 To Len(strMascara)
    If Mid(strMascara, intQtdCaracterMasc, 1) = "#" Then
      strRetorno = strRetorno & Mid(strTexto, intQtdCaracterMasc - intX, 1)
    Else
      'Inserir Máscara
      strRetorno = strRetorno & Mid(strMascara, intQtdCaracterMasc, 1)
      intX = intX + 1
    End If
  Next
  AplicarMascara = strRetorno
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.AplicarMascara]", _
            Err.Description
End Function


Public Function GetUserDataGeral(Bookmark As Variant, _
        Col As Integer, intCOLUNASMATRIZ As Long, intLINHASMATRIZ As Long, mtzMatriz) As Variant
  ' In this example, GetUserData is called by
  ' UnboundReadData to ask the user what data should be
  ' displayed in a specific cell in the grid. The grid
  ' row the cell is in is the one referred to by the
  ' Bookmark parameter, and the column it is in it given
  ' by the Col parameter. GetUserData is called on a
  ' cell-by-cell basis.
  
  On Error GoTo trata
  '
  Dim Index As Long

  ' Figure out which row the bookmark refers to
  Index = IndexFromBookmarkGeral(Bookmark, 0, intLINHASMATRIZ)
  If Index < 0 Or Index >= intLINHASMATRIZ Or _
      Col < 0 Or Col >= intCOLUNASMATRIZ Then
    ' Cell position is invalid, so just return null to
    ' indicate failure
    GetUserDataGeral = Null
  Else
    GetUserDataGeral = mtzMatriz(Col, Index)
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.GetUserDataGeral]", _
            Err.Description
End Function

'Proposito: Retornar o Código do Turno Corrente
Public Function RetornaCodTurnoCorrente(Optional datData As Date) As Long
  On Error GoTo trata
  'Retorna 0 - para Código de Erro
  Dim strSql        As String
  Dim objRs         As ADODB.Recordset
  Dim objGeral      As busSisMed.clsGeral
  Dim lngRetorno    As Long
  '
  Set objGeral = New busSisMed.clsGeral
  '
  strSql = "Select * from Turno Where Status = " & Formata_Dados(True, tpDados_Boolean) & _
    " AND PRONTUARIOID = " & Formata_Dados(giFuncionarioId, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  If objRs.EOF Then
    'Não há turno corrente cadastrado
    lngRetorno = 0
  ElseIf objRs.RecordCount > 1 Then
    'há mais de um turno corrente cadastrado
    lngRetorno = 0
  Else
    lngRetorno = objRs.Fields("PKID").Value
    datData = objRs.Fields("DATA").Value
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  '
  RetornaCodTurnoCorrente = lngRetorno
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.RetornaCodTurnoCorrente]", _
            Err.Description
End Function

Public Sub LerFigurasAvulsasPicBox(ppicBox As PictureBox, pImagem As String, pToolTipText As String)
  On Error GoTo trata
  '
  ppicBox.Picture = LoadPicture(gsIconsPath & pImagem)
  ppicBox.ToolTipText = pToolTipText
  '
  Exit Sub
trata:
  AmpN
  Err.Raise Err.Number, _
            "[mdlUserFunction.LerFigurasAvulsasPicBox]", _
            Err.Description
End Sub

'Proposito: Retornar a descrição do Turno Corrente e emitir msg de erro
'para usuário
Public Function RetornaDescTurnoCorrente(Optional TURNOID As Long) As String
  On Error GoTo trata
  'Retorna "", caso não encontre
  Dim strSql        As String
  Dim objRs         As ADODB.Recordset
  Dim strRetorno    As String
  Dim strErro       As String
  Dim objGeral      As busSisMed.clsGeral
  Dim strDescrTurno As String
  '
  Set objGeral = New busSisMed.clsGeral
  '
  strSql = "Select Data, DIASDASEMANA.DIADASEMANA, Periodo, inicio, termino " & _
    "FROM (PERIODO INNER JOIN TURNO ON PERIODO.PKID = TURNO.PERIODOID) INNER JOIN DIASDASEMANA ON TURNO.DIASDASEMANAID = DIASDASEMANA.PKID " & _
    "WHERE " & IIf(TURNOID <> 0, "TURNO.PKID = " & Formata_Dados(TURNOID, tpDados_Longo) & ";", "Status = " & Formata_Dados(True, tpDados_Boolean) & " AND TURNO.PRONTUARIOID = " & Formata_Dados(giFuncionarioId, tpDados_Longo))

  'ASSUME 0 - TODOS OS DIAS / 1-  FIM DE SEMANA / 2 - FERIADO / 3 - DIAS DE SEMANA / 4 - SEGUNDA / 5 - TERÇA  / 6 - QUARTA / 7 - QUINTA  / 8 - SEXTA / 9 - SÁBADO  / 10 - DOMINGO / 11 - ESPECIAL
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  If objRs.EOF Then
    'Não há turno corrente cadastrado
    strRetorno = ""
    strErro = "Não há Turno aberto no Sistema"
  ElseIf objRs.RecordCount > 1 Then
    'há mais de um turno corrente cadastrado
    strRetorno = ""
    '
    strErro = "Há mais de um turno aberto no sistema:" & vbCrLf & vbCrLf
    Do While Not objRs.EOF
      strDescrTurno = Format(objRs.Fields("Data").Value, "DD/MM/YYYY") & " / " & objRs.Fields("DIADASEMANA").Value & " - Período " & objRs.Fields("Periodo").Value & " de " & objRs.Fields("inicio").Value & " as " & objRs.Fields("termino").Value
      strRetorno = strDescrTurno & vbCrLf
      objRs.MoveNext
    Loop
  Else
    strDescrTurno = Format(objRs.Fields("Data").Value, "DD/MM/YYYY") & " / " & objRs.Fields("DIADASEMANA").Value & " - Período " & objRs.Fields("Periodo").Value & " de " & objRs.Fields("inicio").Value & " as " & objRs.Fields("termino").Value
    strRetorno = strDescrTurno
    strErro = ""
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  '
  RetornaDescTurnoCorrente = strRetorno
  'Emite Msg de Erro
  'If Len(strErro) <> 0 Then Err.Raise 1, , strErro
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.RetornaDescTurnoCorrente]", _
            Err.Description
End Function

Public Function DataHoraAtualFormatada() As String
  On Error GoTo trata
  Dim strRetorno As String
  '
  strRetorno = Format(Now, "DD/MM/YYYY hh:mm")
  '
  DataHoraAtualFormatada = strRetorno
  Exit Function
trata:
  TratarErro Err.Number, Err.Description, "[mdlUserFunction.DataHoraAtualFormatada]"
End Function

'Propósito: Setar o dia da semana
Public Sub SetarDiaDaSemana(objCombo As ComboBox, _
                            dtaCorrente As Date)
  On Error Resume Next
  Dim strRetorno  As String
  '
  Select Case Weekday(dtaCorrente)
  Case vbSunday: strRetorno = "Domingo"
  Case vbMonday: strRetorno = "Segunda"
  Case vbTuesday: strRetorno = "Terça"
  Case vbWednesday: strRetorno = "Quarta"
  Case vbThursday: strRetorno = "Quinta"
  Case vbFriday: strRetorno = "Sexta"
  Case vbSaturday: strRetorno = "Sábado"
  End Select
  '
  objCombo.Text = strRetorno
  Exit Sub
End Sub

Public Sub LimparCampoList(objList As ListBox)
  On Error Resume Next
  objList.Clear
End Sub

Public Sub LimparCampoCombo(objCbo As ComboBox)
  On Error Resume Next
  objCbo.Clear
End Sub
Public Sub LimparCampoTexto(objText As TextBox)
  On Error Resume Next
  objText.Text = ""
End Sub
Public Sub LimparCampoCheck(objCheck As CheckBox)
  On Error Resume Next
  objCheck.Value = False
End Sub
Public Sub LimparCampoOption(objOption As Object)
  On Error Resume Next
  Dim intI As Integer
  For intI = 0 To objOption.Count - 1
    objOption(intI).Value = False
  Next
End Sub
Public Sub LimparCampoMask(objMask As MaskEdBox)
  On Error Resume Next
  Dim strMask As String
  With objMask
    strMask = .Mask
    .Mask = ""
    .Text = ""
    .Mask = strMask
  End With
End Sub
'preencher combo box
Public Sub PreencheCombo(objCbo, _
                         ByVal strSql As String, _
                         Optional TipoTodos As Boolean = True, _
                         Optional TipoBranco As Boolean = False, _
                         Optional SelPriItem As Boolean = False, _
                         Optional strItemSel As String)
  On Error GoTo trata
  Dim objRs         As ADODB.Recordset
  Dim objGeral      As busSisMed.clsGeral
  Dim blnPrimeiro   As Boolean
  Dim strPriItem    As String
  '
  Set objGeral = New busSisMed.clsGeral
  '
  blnPrimeiro = True
  Set objRs = objGeral.ExecutarSQL(strSql)
  objCbo.Clear
  If TipoBranco Then _
     objCbo.AddItem ""
  If TipoTodos Then
     objCbo.AddItem "<TODOS>"
     strPriItem = "<TODOS>"
     blnPrimeiro = False
  End If
  Do While Not objRs.EOF
    objCbo.AddItem objRs.Fields(0) & ""
    If blnPrimeiro Then strPriItem = objRs.Fields(0) & ""
    blnPrimeiro = False
    objRs.MoveNext
  Loop
  If SelPriItem And strPriItem <> "" Then objCbo.Text = strPriItem
  If strItemSel & "" <> "" Then objCbo.Text = strItemSel
  Set objGeral = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.PreencheCombo]", _
            Err.Description
End Sub

Public Function Valida_Moeda(objTarget As Object, _
                             pTipo As TpObriga, _
                             Optional blnSetarFocoControle As Boolean = False, _
                             Optional blnAceitarNegativo As Boolean = True, _
                             Optional blnPintarControle As Boolean = True, _
                             Optional blnValidarPeloClip As Boolean = True) As Boolean
  On Error GoTo trata
  Dim EValor As Boolean
  EValor = True
  'Verifica se Mmaskedit é valor
  If IsNumeric(objTarget.Text) Then
    'é Número, verifica se positivo
    If Not blnAceitarNegativo Then
      If CCur(objTarget.Text) < 0 Then
        'Negativo
        EValor = False
      End If
    End If
  Else
    'Não é número
    EValor = False
  End If
  If EValor Then
  End If
  If pTipo = TpObrigatorio And Not (EValor) Then
    'Campo é obrigatório e não é Valor
    Valida_Moeda = False
  ElseIf pTipo = TpNaoObrigatorio Then
    'Campo não é obrigatório
    If blnValidarPeloClip Then
      If Len(objTarget.ClipText) <> 0 And Not EValor Then
        'Digitou algo que não é Valor
        Valida_Moeda = False
      Else
        Valida_Moeda = True
      End If
    Else
      If Len(objTarget.Text) <> 0 And Not EValor Then
        'Digitou algo que não é Valor
        Valida_Moeda = False
      Else
        Valida_Moeda = True
      End If
    End If
  Else
    Valida_Moeda = True
  End If
  If Valida_Moeda = False Then
    If blnPintarControle Then
      Pintar_Controle objTarget, tpCorContr_Erro
    End If
    If blnSetarFocoControle Then
      SetarFoco objTarget
      blnSetarFocoControle = False
    End If
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.Valida_Moeda]", _
            Err.Description
End Function

Public Function Valida_Option(objOption As Object, _
                              Optional blnSetarFocoControle As Boolean = False) As Boolean
  On Error GoTo trata
  Dim blnRetorno  As Boolean
  Dim intI        As Integer
  blnRetorno = False
  'Verifica se Selecionou um option
  For intI = 0 To objOption.Count - 1
    If objOption(intI).Value = True Then
      blnRetorno = True
      Exit For
    End If
  Next
  If blnRetorno = False Then
    If blnSetarFocoControle Then
      SetarFoco objOption(0)
      blnSetarFocoControle = False
    End If
  End If
  Valida_Option = blnRetorno
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.Valida_Option]", _
            Err.Description
End Function


Public Function Valida_Data(pMsk As MaskEdBox, _
                            pTipo As TpObriga, _
                            Optional blnSetarFocoControle As Boolean = False) As Boolean
  On Error GoTo trata
  Dim EData As Boolean
  EData = True
  'Verifica se Maskedit é data
  If Len(pMsk.Text) = 7 Then
    'Data no formato MM/YYYY
    If Not IsDate("01/" & pMsk.Text) Then
      EData = False
    Else
      If CInt(Mid(pMsk.ClipText, 1, 2)) > 12 Then
        EData = False
      End If
    End If
  ElseIf Len(pMsk.Text) = 5 Then
    'Data no formato DD/MM
    If Not IsDate(pMsk.Text & "/" & Year(Date)) Then
      EData = False
    Else
      If CInt(Right(pMsk.ClipText, 2)) > 12 Then
        EData = False
      End If
    End If
  Else
    If Not IsDate(pMsk.Text) Then
      EData = False
    Else
      If CInt(Mid(pMsk.ClipText, 3, 2)) > 12 Then
        EData = False
      End If
    End If
  End If
  If pTipo = TpObrigatorio And Not (EData) Then
    'Campo é obrigatório e não é data
    Valida_Data = False
  ElseIf pTipo = TpNaoObrigatorio Then
    'Campo não é obrigatório
    If Len(pMsk.ClipText) <> 0 And Not EData Then
      'Digitou algo que não é data
      Valida_Data = False
    Else
      Valida_Data = True
    End If
  Else
    Valida_Data = True
  End If
  If Valida_Data = False Then
    Pintar_Controle pMsk, tpCorContr_Erro
    If blnSetarFocoControle Then
      SetarFoco pMsk
      blnSetarFocoControle = False
    End If
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.Valida_Data]", _
            Err.Description
End Function

Public Function Valida_Data_Str(pData As String, _
                                pTipo As TpObriga) As Boolean
  On Error GoTo trata
  Dim EData As Boolean
  EData = True
  'Verifica se Maskedit é data
  If Len(pData) = 7 Then
'''    'Data no formato MM/YYYY
'''    If Not IsDate("01/" & pData.Text) Then
'''      EData = False
'''    Else
'''      If CInt(Mid(pData.ClipText, 1, 2)) > 12 Then
'''        EData = False
'''      End If
'''    End If
'''  ElseIf Len(pData.Text) = 5 Then
'''    'Data no formato DD/MM
'''    If Not IsDate(pData.Text & "/" & Year(Date)) Then
'''      EData = False
'''    Else
'''      If CInt(Right(pData.ClipText, 2)) > 12 Then
'''        EData = False
'''      End If
'''    End If
  Else
    If Not IsDate(pData) Then
      EData = False
    Else
      If CInt(Mid(pData, 4, 2)) > 12 Then
        EData = False
      End If
    End If
  End If
  If pTipo = TpObrigatorio And Not (EData) Then
    'Campo é obrigatório e não é data
    Valida_Data_Str = False
  ElseIf pTipo = TpNaoObrigatorio Then
    'Campo não é obrigatório
    If Len(pData) <> 0 And Not EData Then
      'Digitou algo que não é data
      Valida_Data_Str = False
    Else
      Valida_Data_Str = True
    End If
  Else
    Valida_Data_Str = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.Valida_Data_Str]", _
            Err.Description
End Function



Public Function TestaCPF(CPF As String) As Boolean
  'Recebe o CPF e informa se é falso ou verdadeiro
  On Error GoTo trata
  Dim CPF1      As String
  Dim CPF2      As String
  Dim Soma      As Integer
  Dim Digito    As String
  Dim I         As Integer
  Dim J         As Integer
  Dim Controle  As String
  Dim Mult      As String
  '
  Dim Resto, Digito1, Digito2
  'Identifica as duas partes do CPF
  CPF1 = Left$(CPF, 10)
  CPF2 = Right$(CPF, 2)
  
  'Multiplicadores que fazem parte do algorítimo de checagem
  Mult = "1098765432"
  
  'Inicializa a variável de controle
  Controle = ""
  
  'Loop de verificação
  'Calculo do primeiro digito verificador
  
  Soma = 0
  I = 1
  Soma = Soma + (Val(Mid$(CPF1, I, 1)) * Val(Mid$(Mult, I, 2)))
  For I = 2 To 9
      If Mid$(CPF1, I, 1) <> "/" Then
       Soma = Soma + (Val(Mid$(CPF1, I, 1)) * Val(Mid$(Mult, I + 1, 1)))
      End If
  Next I
  
  Resto = Soma Mod 11
      
  If Resto = 1 Or Resto = 0 Then
      Digito1 = 0
  Else
      Digito1 = 11 - Resto
  End If
      
    
  'Sequência de multiplicadores para o cáculo so segundo dígito
   Mult = "11109876543"
   
  'Loop de verificação
  'Calculo do segundo digito verificador
  
  Soma = 0
  
  I = 1
  Soma = Soma + (Val(Mid$(CPF1, I, 1)) * Val(Mid$(Mult, I, 2)))
  I = I + 1
  Soma = Soma + (Val(Mid$(CPF1, I, 1)) * Val(Mid$(Mult, I + 1, 2)))
  
  For I = 3 To 9
      If Mid$(CPF1, I, 1) <> "/" Then
          Soma = Soma + (Val(Mid$(CPF1, I, 1)) * Val(Mid$(Mult, I + 2, 1)))
      End If
  Next I
  
  Soma = Soma + (Digito1 * 2)
  
  Resto = Soma Mod 11
     
  If Resto = 1 Or Resto = 0 Then
      Digito2 = 0
  Else
      Digito2 = 11 - Resto
  End If
  
  
  'Compara os dígitos calculados (COntrole) com dígitos informados (CGC2)
  Digito = Digito1 & Digito2
  Controle = Controle + Trim$(CStr(Digito))
  
  If Controle <> CPF2 Then
      TestaCPF = False
  Else
      TestaCPF = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.TestaCPF]", _
            Err.Description
End Function
'Proposito: Retornar o dia da semana da data corrente
'   OU DA DATA DE ENTRADA recebida como parâmetro
Public Function Retorna_DIADASEMANA_Data(Optional dtDATAENTRADA As Date) As Integer
  'Retorna -1 - para Código de Erro
  Dim strSql          As String
  Dim objRs           As ADODB.Recordset
  Dim intRetorno      As Integer
  Dim objGeral        As busSisMed.clsGeral
  Dim datDataBase     As Date
  Dim blnFeriado      As Boolean
  '
  On Error GoTo trata
  Set objGeral = New busSisMed.clsGeral
  '
  intRetorno = -1
  If IsDataValida(dtDATAENTRADA) = False Then
    datDataBase = Now
  Else
    datDataBase = dtDATAENTRADA
  End If
  '------------------------------------
  'Verificação se é feriado
  blnFeriado = False
'''  strSql = "SELECT * FROM FERIADO WHERE DIAMES = " & Formata_Dados(Format(datDataBase, "DD/MM"), tpDados_Texto)
'''  Set objRs = objGeral.ExecutarSQL(strSql)
'''  If Not objRs.EOF Then
'''    blnFeriado = True
'''    'Captura feriado
'''    intRetorno = objRs.Fields("DIASDASEMANAID").Value
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
  'fim da verificação se é feriado
  '------------------------------------
  
  '------------------------------------
  'Obter dia da semana caso não seja feriado
  If blnFeriado = False Then
    strSql = "SELECT PKID FROM DIASDASEMANA WHERE UPPER(DIADASEMANA) = " & Formata_Dados(Retorna_DIADASEMANA_Descr(datDataBase), tpDados_Texto)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      'Captura dia da semana
      intRetorno = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  'Fim de Obter dia da semana caso não seja feriado
  '------------------------------------
  Set objGeral = Nothing
  '
  Retorna_DIADASEMANA_Data = intRetorno
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.Retorna_DIADASEMANA_Data]", _
            Err.Description
End Function

'Proposito: Retornar o dia da semana da data corrente
'   OU DA DATA DE ENTRADA recebida como parâmetro
Public Function Retorna_CAIXA_Nome(lngPRESTADORID As Long) As String
  'Retorna -1 - para Código de Erro
  Dim strSql          As String
  Dim objRs           As ADODB.Recordset
  Dim strRetorno      As String
  Dim objGeral        As busSisMed.clsGeral
  Dim datDataBase     As Date
  '
  On Error GoTo trata
  Set objGeral = New busSisMed.clsGeral
  '
  strRetorno = -1
  '------------------------------------
  'Obter dia do funcionário
  strSql = "SELECT PRONTUARIO.NOME FROM PRONTUARIO INNER JOIN FUNCIONARIO ON PRONTUARIO.PKID = FUNCIONARIO.PRONTUARIOID WHERE PRONTUARIO.PKID = " & Formata_Dados(lngPRESTADORID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    'Captura dia da semana
    strRetorno = objRs.Fields("NOME").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'Fim de Obter dia da semana caso não seja feriado
  '------------------------------------
  Set objGeral = Nothing
  '
  Retorna_CAIXA_Nome = strRetorno
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.Retorna_CAIXA_Nome]", _
            Err.Description
End Function


Public Function IsDataValida(dtData As Date) As Boolean
  On Error GoTo trata
  Dim dtInvalida        As Date
  Dim blnIsDataValida   As Boolean
  dtInvalida = CDate("30/12/1899 00:00:00")
  blnIsDataValida = True
  If dtData = dtInvalida Then
    blnIsDataValida = False
  End If
  IsDataValida = blnIsDataValida
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.IsDataValida]", _
            Err.Description
End Function

Public Function Retorna_DIADASEMANA_Descr(dtDATAENTRADA As Date) As String
  On Error GoTo trata
  Dim strDiaSemana As String
  Select Case Weekday(dtDATAENTRADA)
  Case 1: strDiaSemana = "DOMINGO"
  Case 2: strDiaSemana = "SEGUNDA"
  Case 3: strDiaSemana = "TERÇA"
  Case 4: strDiaSemana = "QUARTA"
  Case 5: strDiaSemana = "QUINTA"
  Case 6: strDiaSemana = "SEXTA"
  Case 7: strDiaSemana = "SÁBADO"
  Case Else: strDiaSemana = ""
  End Select
  Retorna_DIADASEMANA_Descr = strDiaSemana
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.Retorna_DIADASEMANA_Descr]", _
            Err.Description
End Function

Public Function TRANSFORMA_MAIUSCULA(pKeyAscii) As String
  On Error GoTo trata
  TRANSFORMA_MAIUSCULA = Asc(UCase(Chr(pKeyAscii)))
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.TRANSFORMA_MAIUSCULA]", _
            Err.Description
End Function



'Propósito: Retornar e Gravar o Sequencial geral
Public Function RetornaGravaCampoSequencial(strCampo As String, _
                                            lngTURNOID As Long) As Long
  On Error GoTo trata
  '
  Dim objGeral  As busSisMed.clsGeral
  Dim objRs     As ADODB.Recordset
  Dim strSql    As String
  Dim lngSeq    As Long
  '
  Set objGeral = New busSisMed.clsGeral
  strSql = "SELECT ISNULL(MAX(" & strCampo & "), 0) + 1 AS SEQ " & _
    "From SEQUENCIAL " & _
    " WHERE TURNOID = " & Formata_Dados(lngTURNOID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  If objRs.EOF Then
    lngSeq = 1
  ElseIf Not IsNumeric(objRs.Fields("SEQ").Value) Then
    lngSeq = 1
  Else
    lngSeq = objRs.Fields("SEQ").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  If lngSeq = 1 Then
    strSql = "INSERT INTO SEQUENCIAL (" & strCampo & ", TURNOID) VALUES (" & _
        Formata_Dados(lngSeq, tpDados_Longo) & _
        ", " & Formata_Dados(lngTURNOID, tpDados_Longo) & _
        ")"
  Else
    strSql = "UPDATE SEQUENCIAL SET " & strCampo & " = " & _
      Formata_Dados(lngSeq, tpDados_Longo) & _
      " WHERE TURNOID = " & Formata_Dados(lngTURNOID, tpDados_Longo)
  End If
  '---
  objGeral.ExecutarSQLAtualizacao (strSql)
  '
  RetornaGravaCampoSequencial = CLng(lngSeq)
  '
  Set objGeral = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, "[mdlUserFunction.RetornaGravaCampoSequencial]", Err.Description
End Function

'Propósito: Retornar e Gravar o Sequencial (senha)
Public Function RetornaGravaCampoSequencialSenha(strCampo As String, _
                                                 lngPRESTADORID As Long) As Long
  On Error GoTo trata
  '
  Dim objGeral        As busSisMed.clsGeral
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  Dim lngSeq          As Long
  Dim strDataAtual    As String
  '
  strDataAtual = Format(Now, "DD/MM/YYYY")
'''  Set objGeral = New busSisMed.clsGeral
'''  strSql = "SELECT ISNULL(MAX(" & strCampo & "), 0) + 1 AS SEQ " & _
'''    "From SEQUENCIALPRESTADOR " & _
'''    " WHERE PRESTADORID = " & Formata_Dados(lngPRESTADORID, tpDados_Longo) & _
'''    " AND DATA = " & Formata_Dados(strDataAtual, tpDados_DataHora)
'''  Set objRs = objGeral.ExecutarSQL(strSql)
'''  '
'''  If objRs.EOF Then
'''    lngSeq = 1
'''  ElseIf Not IsNumeric(objRs.Fields("SEQ").Value) Then
'''    lngSeq = 1
'''  Else
'''    lngSeq = objRs.Fields("SEQ").Value
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
'''  '
'''  If lngSeq = 1 Then
'''    strSql = "INSERT INTO SEQUENCIALPRESTADOR (" & strCampo & ", DATA, PRESTADORID) VALUES (" & _
'''        Formata_Dados(lngSeq, tpDados_Longo) & _
'''        ", " & Formata_Dados(strDataAtual, tpDados_DataHora) & _
'''        ", " & Formata_Dados(lngPRESTADORID, tpDados_Longo) & _
'''        ")"
'''  Else
'''    strSql = "UPDATE SEQUENCIALPRESTADOR SET " & strCampo & " = " & _
'''      Formata_Dados(lngSeq, tpDados_Longo) & _
'''      " WHERE PRESTADORID = " & Formata_Dados(lngPRESTADORID, tpDados_Longo) & _
'''      " AND DATA = " & Formata_Dados(strDataAtual, tpDados_DataHora)
'''  End If
'''  '---
'''  objGeral.ExecutarSQLAtualizacao (strSql)
  lngSeq = 0
  Set objGeral = New busSisMed.clsGeral
  'strSql = "EXEC SP_RETORNAGRAVASEQUENCIALSENHA " & lngPRESTADORID & ", '" & strDataAtual & "'"
  lngSeq = objGeral.ExecutarSQLRetInteger("SP_RETORNAGRAVASEQUENCIALSENHA", Array( _
                                                    mp("@PRESTADORID", adInteger, 4, lngPRESTADORID), _
                                                    mp("@DATA", adVarChar, 30, strDataAtual)))
  '
  RetornaGravaCampoSequencialSenha = CLng(lngSeq)
  '
  Set objGeral = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, "[mdlUserFunction.RetornaGravaCampoSequencialSenha]", Err.Description
End Function

Public Sub LerFigurasAvulsas(pbtn As CommandButton, pImagem As String, pImagemDown As String, pToolTipText As String)
  On Error GoTo trata
  '
  pbtn.Picture = LoadPicture(gsIconsPath & pImagem)
  pbtn.DownPicture = LoadPicture(gsIconsPath & pImagemDown)
  pbtn.ToolTipText = pToolTipText
  '
  Exit Sub
trata:
  AmpN
  Err.Raise Err.Number, _
            "[mdlUserFunction.LerFigurasAvulsas]", _
            Err.Description
End Sub



'Proposito: Verificar movimento após o fechamento
Public Sub VerificaMovAposFecha(lngGRID As Long)
  On Error GoTo trata
  'Retorna 0 - para Código de Erro
  Dim strSql        As String
  Dim objRs         As ADODB.Recordset
  Dim objGeral      As busSisMed.clsGeral
  Dim vrValorJaPago   As Currency
  Dim vrTotGR         As Currency
  Dim objGR           As busSisMed.clsGR
  Dim strStatus       As String
  '
  Set objGeral = New busSisMed.clsGeral
  '
  'Capturar valor total da GR
  vrTotGR = 0
  strSql = "SELECT SUM(VALOR) AS VALOR, MIN(GR.STATUS) AS STATUS " & _
    "FROM GRPROCEDIMENTO INNER JOIN GR ON GR.PKID = GRPROCEDIMENTO.GRID " & _
    "WHERE GRID = " & Formata_Dados(lngGRID, tpDados_Longo, tpNulo_Aceita)

  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    If Not IsNull(objRs.Fields("VALOR").Value) Then
      vrTotGR = objRs.Fields("VALOR").Value
    End If
    strStatus = objRs.Fields("STATUS").Value & ""
  End If
  objRs.Close
  Set objRs = Nothing
  'Sai se GR ainda não estiver fechada
  If strStatus <> "F" Then Exit Sub
  'Capturar valor já pago
  vrValorJaPago = 0
  strSql = "SELECT SUM(VALOR * (CASE INDDEBITOCREDITO WHEN 'C' THEN 1 ELSE -1 END)) AS VALORJAPAGO, SUM(VRGORJETA) AS VRGORJETAJAPAGO, SUM(VRTROCO) AS VRTROCOJAPAGO " & _
    "FROM CONTACORRENTE " & _
    "WHERE GRID = " & Formata_Dados(lngGRID, tpDados_Longo, tpNulo_Aceita)

  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    If Not IsNull(objRs.Fields("VALORJAPAGO").Value) Then
      vrValorJaPago = objRs.Fields("VALORJAPAGO").Value
    End If
    If Not IsNull(objRs.Fields("VRGORJETAJAPAGO").Value) Then
      vrValorJaPago = vrValorJaPago - objRs.Fields("VRGORJETAJAPAGO").Value
    End If
    If Not IsNull(objRs.Fields("VRTROCOJAPAGO").Value) Then
      vrValorJaPago = vrValorJaPago - objRs.Fields("VRTROCOJAPAGO").Value
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  If vrValorJaPago <> vrTotGR Then
    'Valor do pagamento < que valor a pagar
    Set objGR = New busSisMed.clsGR
    objGR.AlterarStatusGR lngGRID, _
                          "M", _
                          ""
    Set objGR = Nothing
  End If
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.VerificaMovAposFecha]", _
            Err.Description
End Sub

'Proposito: Retornar o Código do Turno Corrente
Public Function RetornaCodTurnosPelaData(datData As Date, _
                                         strNivel As String, _
                                         lngFuncionarioId As Long) As String
  On Error GoTo trata
  'Retorna 0 - para Código de Erro
  Dim strSql        As String
  Dim objRs         As ADODB.Recordset
  Dim objGeral      As busSisMed.clsGeral
  Dim lngRetorno    As Long
  Dim strData       As String
  Dim datDataSeg    As Date
  Dim strDataSeg    As String
  Dim strRetorno    As String
  '
  Set objGeral = New busSisMed.clsGeral
  '
  datDataSeg = DateAdd("d", 1, datData)
  strData = Format(datData, "DD/MM/YYYY") & " 00:00"
  strDataSeg = Format(datDataSeg, "DD/MM/YYYY") & " 23:59"
  
  strSql = "SELECT TURNO.PKID FROM TURNO " & _
            " INNER JOIN FUNCIONARIO ON FUNCIONARIO.PRONTUARIOID = TURNO.PRONTUARIOID " & _
            " WHERE TURNO.DATA >= " & Formata_Dados(strData, tpDados_DataHora) & _
            " AND TURNO.DATA < " & Formata_Dados(strDataSeg, tpDados_DataHora)
  If strNivel = gsLaboratorio Then
    strSql = strSql & " AND TURNO.PRONTUARIOID = " & Formata_Dados(lngFuncionarioId, tpDados_Longo)
  Else
    strSql = strSql & " AND (TURNO.PRONTUARIOID = " & Formata_Dados(lngFuncionarioId, tpDados_Longo) & _
            " OR FUNCIONARIO.NIVEL = " & Formata_Dados(gsLaboratorio, tpDados_Texto) & ")"
  End If
  
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  strRetorno = "("
  If objRs.EOF Then
    'Não há turno corrente cadastrado
    strRetorno = strRetorno & "0"
  Else
    Do While Not objRs.EOF
      If strRetorno <> "(" Then strRetorno = strRetorno & ","
      strRetorno = strRetorno & objRs.Fields("PKID").Value
      objRs.MoveNext
    Loop
  End If
  strRetorno = strRetorno & ")"
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  '
  RetornaCodTurnosPelaData = strRetorno
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.RetornaCodTurnosPelaData]", _
            Err.Description
End Function
'Proposito: Retornar o Código do Turno Corrente
Public Function RetornaCodTurnosPelaDataTODOS(datData As Date) As String
  On Error GoTo trata
  'Retorna 0 - para Código de Erro
  Dim strSql        As String
  Dim objRs         As ADODB.Recordset
  Dim objGeral      As busSisMed.clsGeral
  Dim lngRetorno    As Long
  Dim strData       As String
  Dim datDataSeg    As Date
  Dim strDataSeg    As String
  Dim strRetorno    As String
  '
  Set objGeral = New busSisMed.clsGeral
  '
  datDataSeg = DateAdd("d", 1, datData)
  strData = Format(datData, "DD/MM/YYYY") & " 00:00"
  strDataSeg = Format(datDataSeg, "DD/MM/YYYY") & " 23:59"
  
  strSql = "SELECT TURNO.PKID FROM TURNO " & _
            " INNER JOIN FUNCIONARIO ON FUNCIONARIO.PRONTUARIOID = TURNO.PRONTUARIOID " & _
            " WHERE TURNO.DATA >= " & Formata_Dados(strData, tpDados_DataHora) & _
            " AND TURNO.DATA < " & Formata_Dados(strDataSeg, tpDados_DataHora)
  
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  strRetorno = "("
  If objRs.EOF Then
    'Não há turno corrente cadastrado
    strRetorno = strRetorno & "0"
  Else
    Do While Not objRs.EOF
      If strRetorno <> "(" Then strRetorno = strRetorno & ","
      strRetorno = strRetorno & objRs.Fields("PKID").Value
      objRs.MoveNext
    Loop
  End If
  strRetorno = strRetorno & ")"
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  '
  RetornaCodTurnosPelaDataTODOS = strRetorno
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.RetornaCodTurnosPelaDataTODOS]", _
            Err.Description
End Function



'Propósito: Carregar combos com período
Public Sub MovFinanc(vrCaixaInicial As Currency, _
                     vrCheque As Currency, _
                     vrEspecie As Currency, _
                     vrCartao As Currency, _
                     vrCartaoDeb As Currency, _
                     vrPenhor As Currency, _
                     vrFatura As Currency, _
                     vrGorjeta As Currency, _
                     vrSaldoDespDiveDin As Currency, _
                     vrSaldoDespDiveChqResg As Currency, _
                     vrTroco As Currency, _
                     vrDepDin As Currency, _
                     vrDepChq As Currency, _
                     vrDepCar As Currency, _
                     vrDepCarDeb As Currency, _
                     vrDepPen As Currency, _
                     vrDepFat As Currency, _
                     vrRetDin As Currency, _
                     vrRetChq As Currency, _
                     vrRetCar As Currency, _
                     vrRetCarDeb As Currency, _
                     vrRetPen As Currency, _
                     vrRetFat As Currency, _
                     vrTotComMov As Currency)
  On Error GoTo trata
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim objGeral  As busSisMed.clsGeral
  '
  Set objGeral = New busSisMed.clsGeral
  '--- Pega Valor Inicial do Caixa
  strSql = "Select * From turno Where pkid = " & RetornaCodTurnoCorrente
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  If objRs.EOF Then
    vrCaixaInicial = 0
  ElseIf Not IsNumeric(objRs.Fields("vrCaixaInicial").Value) Then
    vrCaixaInicial = 0
  Else
    vrCaixaInicial = CCur(objRs.Fields("vrCaixaInicial").Value)
  End If
  objRs.Close
  Set objRs = Nothing
  '
  '---- Pega valores somados na locação
  'strSql = "Select SUM(locacao.PGTOCARTAODEB) as PGTOCARTAODEB, SUM(locacao.VRCALCTROCO) as PGTOTROCO, SUM(locacao.PGTOESPECIE) as PGTOESPECIE, SUM(locacao.PGTOCHEQUE) as PGTOCHEQUE, SUM(locacao.PGTOCARTAO) as PGTOCARTAO, SUM(locacao.VRCALCGORJETA) as PGTOGORJETA, SUM(locacao.PGTOPENHOR) as PGTOPENHOR From LOCACAO INNER JOIN TURNO ON (TURNO.PKID = LOCACAO.TURNORECEBEID) Where TURNO.pkid = " & RetornaCodTurnoCorrente
  strSql = "Select " & _
    " PGTOCARTAODEB " & _
    ", PGTOTROCO " & _
    ", PGTOGORJETA " & _
    ", PGTOESPECIE " & _
    ", PGTOCHEQUE " & _
    ", PGTOCARTAO " & _
    ", PGTOPENHOR " & _
    ", PGTOFATURA " & _
    " FROM vw_cons_t_cred  " & _
    " WHERE vw_cons_t_cred.TURNOID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo)
    
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  If objRs.EOF Then
    vrCheque = 0
    vrEspecie = 0
    vrCartao = 0
    vrCartaoDeb = 0
    vrPenhor = 0
    vrFatura = 0
    vrGorjeta = 0
    vrTroco = 0
    vrSaldoDespDiveDin = 0
    vrSaldoDespDiveChqResg = 0
  Else
    vrCheque = IIf(Not IsNumeric(objRs.Fields("PGTOCHEQUE").Value), 0, objRs.Fields("PGTOCHEQUE").Value)
    vrEspecie = IIf(Not IsNumeric(objRs.Fields("PGTOESPECIE").Value), 0, objRs.Fields("PGTOESPECIE").Value)
    vrCartao = IIf(Not IsNumeric(objRs.Fields("PGTOCARTAO").Value), 0, objRs.Fields("PGTOCARTAO").Value)
    vrCartaoDeb = IIf(Not IsNumeric(objRs.Fields("PGTOCARTAODEB").Value), 0, objRs.Fields("PGTOCARTAODEB").Value)
    vrPenhor = IIf(Not IsNumeric(objRs.Fields("PGTOPENHOR").Value), 0, objRs.Fields("PGTOPENHOR").Value)
    vrFatura = IIf(Not IsNumeric(objRs.Fields("PGTOFATURA").Value), 0, objRs.Fields("PGTOFATURA").Value)
    vrGorjeta = IIf(Not IsNumeric(objRs.Fields("PGTOGORJETA").Value), 0, objRs.Fields("PGTOGORJETA").Value)
    vrTroco = IIf(Not IsNumeric(objRs.Fields("PGTOTROCO").Value), 0, objRs.Fields("PGTOTROCO").Value)
    vrSaldoDespDiveDin = 0
    vrSaldoDespDiveChqResg = 0
  End If
  objRs.Close
  Set objRs = Nothing
  '
  '---
  'NOVO - Calcular totais com movimentações no turno
  strSql = "Select SUM(VRDEPPEN) AS VRDEPPEN1, SUM(VRDEPFAT) AS VRDEPFAT1, SUM(VRRETPEN) AS VRRETPEN1, SUM(VRRETFAT) AS VRRETFAT1, SUM(VRRETDIN) AS VRRETDIN1, SUM(VRRETCHQ) AS VRRETCHQ1, SUM(VRRETCAR) AS VRRETCAR1, SUM(VRRETCARDEB) AS VRRETCARDEB1, SUM(VRDEPDIN) AS VRDEPDIN1, SUM(VRDEPCHQ) AS VRDEPCHQ1, SUM(VRDEPCAR) AS VRDEPCAR1, SUM(VRDEPCARDEB) AS VRDEPCARDEB1 FROM SANGRIA " & _
   "Where TURNOID = " & RetornaCodTurnoCorrente & ";"
  'strSql = "Select SUM(TAB_VENDAITEM.VALOR) as VALOR From DESPESA INNER JOIN TURNO ON (TURNO.PKID = DESPESA.TURNOID) Where TURNO.pkid = " & RetornaCodTurnoCorrente
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  If objRs.EOF Then
    vrDepDin = 0
    vrDepChq = 0
    vrDepCar = 0
    vrDepCarDeb = 0
    vrDepPen = 0
    vrDepFat = 0
    vrRetDin = 0
    vrRetChq = 0
    vrRetCar = 0
    vrRetCarDeb = 0
    vrRetPen = 0
    vrRetFat = 0
  Else
    vrDepDin = IIf(Not IsNumeric(objRs.Fields("vrDepDin1").Value), 0, objRs.Fields("vrDepDin1").Value)
    vrDepChq = IIf(Not IsNumeric(objRs.Fields("vrDepChq1").Value), 0, objRs.Fields("vrDepChq1").Value)
    vrDepCar = IIf(Not IsNumeric(objRs.Fields("vrDepCar1").Value), 0, objRs.Fields("vrDepCar1").Value)
    vrDepCarDeb = IIf(Not IsNumeric(objRs.Fields("vrDepCarDeb1").Value), 0, objRs.Fields("vrDepCarDeb1").Value)
    vrDepPen = IIf(Not IsNumeric(objRs.Fields("vrDepPen1").Value), 0, objRs.Fields("vrDepPen1").Value)
    vrDepFat = IIf(Not IsNumeric(objRs.Fields("vrDepFat1").Value), 0, objRs.Fields("vrDepFat1").Value)
    vrRetDin = IIf(Not IsNumeric(objRs.Fields("vrRetDin1").Value), 0, objRs.Fields("vrRetDin1").Value)
    vrRetChq = IIf(Not IsNumeric(objRs.Fields("vrRetChq1").Value), 0, objRs.Fields("vrRetChq1").Value)
    vrRetCar = IIf(Not IsNumeric(objRs.Fields("vrRetCar1").Value), 0, objRs.Fields("vrRetCar1").Value)
    vrRetCarDeb = IIf(Not IsNumeric(objRs.Fields("vrRetCarDeb1").Value), 0, objRs.Fields("vrRetCarDeb1").Value)
    vrRetPen = IIf(Not IsNumeric(objRs.Fields("vrRetPen1").Value), 0, objRs.Fields("vrRetPen1").Value)
    vrRetFat = IIf(Not IsNumeric(objRs.Fields("vrRetFat1").Value), 0, objRs.Fields("vrRetFat1").Value)
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'Pega total
  vrTotComMov = vrEspecie + vrCheque + vrCartao + vrCartaoDeb + vrPenhor + vrFatura - vrSaldoDespDiveChqResg - vrSaldoDespDiveDin - vrTroco
  'Soma com Movimentação do caixa
  vrTotComMov = vrTotComMov + vrDepPen + vrDepFat + vrDepDin + vrDepChq + vrDepCar + vrDepCarDeb - vrRetDin - vrRetChq - vrRetCar - vrRetCarDeb - vrRetPen - vrRetFat
  '
  Set objGeral = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.MovFinanc]", _
            Err.Description
End Sub

Public Function RetornaDescAtende(strAtendente As String) As String
  On Error GoTo trata
  Dim strRetorno As String
  strRetorno = strAtendente
  strRetorno = Replace(strRetorno, "(" & gsAdmin & ") ", "")
  strRetorno = Replace(strRetorno, "(" & gsDiretor & ") ", "")
  strRetorno = Replace(strRetorno, "(" & gsGerente & ") ", "")
  strRetorno = Replace(strRetorno, "(" & gsCaixa & ") ", "")
  strRetorno = Replace(strRetorno, "(" & gsLaboratorio & ") ", "")
  strRetorno = Replace(strRetorno, "(" & gsFinanceiro & ") ", "")
  '
  RetornaDescAtende = strRetorno
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.RetornaDescAtende]", _
            Err.Description
End Function

Public Function RetornaNivelAtende(strAtendente As String) As String
  On Error GoTo trata
  Dim strRetorno As String
  strRetorno = strAtendente
  strRetorno = Mid(strRetorno, 2, 3)
  '
  RetornaNivelAtende = strRetorno
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.RetornaNivelAtende]", _
            Err.Description
End Function

Public Sub RegistrarChave()
  On Error GoTo trata
  Dim objGeral    As busSisMed.clsGeral
  Dim objProtec   As busSisMed.clsProtec
  
  'PEDE SENHA SUPERIOR E GRAVA LOG
  '----------------------------
  '----------------------------
  'Pede Senha Superior (Diretor, Gerente ou Administrador
  If gsNivel <> "ADM" And gsNivel <> "DIR" Then
    'Só pede senha superior se quem estiver logado não for superior
    frmUserLoginSup.Show vbModal
    
    If Len(Trim(gsNomeUsuLib)) = 0 Then
      MsgBox "É necessário a confirmação com senha de administrador para registrar a chave.", vbExclamation, TITULOSISTEMA
      Exit Sub
    End If
    '
    'Capturou Nome do Usuário, continua processo
  Else
    gsNomeUsuLib = gsNomeUsu
  End If
  'If Len(Msg) = 0 Then
    'Inclui Log
  '  INCLUI_LOG_UNIDADE MODOALTERAR, Data1.Recordset!PKID, "Alteração do depósito", "Unidade " & sNumeroAptoPrinc, "", "", "", gsNomeUsuLib
  'End If
  '---------------------------------------------------------------
  '----------------
  'Proteção do sistema
  '----------------
  Set objGeral = New busSisMed.clsGeral
  Set objProtec = New busSisMed.clsProtec
  '----------------
  'Verifica Proteção do sistema
  '-------------------------
  'Valida primeira vez que entrou no sistema
  objProtec.Gravar_Chave objGeral.ObterConnectionString
  '
  Set objProtec = Nothing
  Set objGeral = Nothing
  '-----------------
  '------------ FIM
  '----------------
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.RegistrarChave]", _
            Err.Description
End Sub

Public Function Extenso(ByVal Valor As _
      Double, ByVal MoedaPlural As _
      String, ByVal MoedaSingular As _
      String) As String
  On Error GoTo trata
  Dim StrValor As String, Negativo As Boolean
  Dim Buf As String, Parcial As Integer
  Dim Posicao As Integer, Unidades
  Dim Dezenas, Centenas, PotenciasSingular
  Dim PotenciasPlural
  
  Negativo = (Valor < 0)
  Valor = Abs(CDec(Valor))
  If Valor Then
     Unidades = Array(vbNullString, "Um", "Dois", _
                "Três", "Quatro", "Cinco", _
                "Seis", "Sete", "Oito", "Nove", _
                "Dez", "Onze", "Doze", "Treze", _
                "Quatorze", "Quinze", "Dezesseis", _
                "Dezessete", "Dezoito", "Dezenove")
     Dezenas = Array(vbNullString, vbNullString, _
               "Vinte", "Trinta", "Quarenta", _
               "Cinqüenta", "Sessenta", "Setenta", _
               "Oitenta", "Noventa")
     Centenas = Array(vbNullString, "Cento", _
                "Duzentos", "Trezentos", _
                "Quatrocentos", "Quinhentos", _
                "Seiscentos", "Setecentos", _
                "Oitocentos", "Novecentos")
     PotenciasSingular = Array(vbNullString, " Mil", _
                         " Milhão", " Bilhão", _
                         " Trilhão", " Quatrilhão")
     PotenciasPlural = Array(vbNullString, " Mil", _
                       " Milhões", " Bilhões", _
                       " Trilhões", " Quatrilhões")
  
     StrValor = Left(Format(Valor, String(18, "0") & _
                ".000"), 18)
     For Posicao = 1 To 18 Step 3
       Parcial = Val(Mid(StrValor, Posicao, 3))
       If Parcial Then
         If Parcial = 1 Then
           Buf = "Um" & PotenciasSingular((18 - _
                 Posicao) \ 3)
         ElseIf Parcial = 100 Then
           Buf = "Cem" & PotenciasSingular((18 - _
                 Posicao) \ 3)
         Else
           Buf = Centenas(Parcial \ 100)
           Parcial = Parcial Mod 100
           If Parcial <> 0 And Buf <> vbNullString Then
             Buf = Buf & " e "
           End If
           If Parcial < 20 Then
             Buf = Buf & Unidades(Parcial)
           Else
             Buf = Buf & Dezenas(Parcial \ 10)
             Parcial = Parcial Mod 10
             If Parcial <> 0 And Buf <> vbNullString Then
               Buf = Buf & " e "
             End If
             Buf = Buf & Unidades(Parcial)
           End If
           Buf = Buf & PotenciasPlural((18 - Posicao) \ 3)
         End If
         If Buf <> vbNullString Then
           If Extenso <> vbNullString Then
             Parcial = Val(Mid(StrValor, Posicao, 3))
             If Posicao = 16 And (Parcial < 100 Or _
                 (Parcial Mod 100) = 0) Then
               Extenso = Extenso & " e "
             Else
               Extenso = Extenso & ", "
             End If
           End If
           Extenso = Extenso & Buf
         End If
       End If
     Next
     If Extenso <> vbNullString Then
       If Negativo Then
         Extenso = "Menos " & Extenso
       End If
       If Int(Valor) = 1 Then
         Extenso = Extenso & " " & MoedaSingular
       Else
         Extenso = Extenso & " " & MoedaPlural
       End If
     End If
     Parcial = Int((Valor - Int(Valor)) * _
               100 + 0.1)
     If Parcial Then
       Buf = Extenso(Parcial, "Centavos", _
             "Centavo")
       If Extenso <> vbNullString Then
         Extenso = Extenso & " e "
       End If
       Extenso = Extenso & Buf
     End If
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.Extenso]", _
            Err.Description

End Function

Public Sub TratarAcValorProco(strAceitaValor As String, _
                              objMaskValor As MaskEdBox, _
                              objQuantidade As MaskEdBox, _
                              Optional blnSetarFoco As Boolean = True)
  On Error GoTo trata
  If strAceitaValor = "S" Then
    objMaskValor.BackColor = -2147483643
    objMaskValor.Enabled = True
    If blnSetarFoco Then _
      SetarFoco objMaskValor
  ElseIf strAceitaValor = "N" Then
    objMaskValor.BackColor = 14737632
    objMaskValor.Enabled = False
    If blnSetarFoco Then _
      SetarFoco objQuantidade
  Else
    objMaskValor.BackColor = 14737632
    objMaskValor.Enabled = False
    If blnSetarFoco Then _
      SetarFoco objQuantidade
  End If
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.TratarAcValorProco]", _
            Err.Description
End Sub
      


Private Function ValidaCamposFechaTurno() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  Dim strSql                As String
  Dim objRs                 As ADODB.Recordset
  Dim objGeral              As busSisMed.clsGeral
  '
  blnSetarFocoControle = True
  ValidaCamposFechaTurno = False
  '
  Set objGeral = New busSisMed.clsGeral
  '
  strSql = "SELECT * FROM GR " & _
      " WHERE TURNOID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo) & _
      " AND (GR.STATUS = " & Formata_Dados("I", tpDados_Texto) & _
      " OR GR.STATUS = " & Formata_Dados("M", tpDados_Texto) & ")"

'STATUS INICIAL OU COM MOVIMENTO APÓS O FECHAMENTO
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    strMsg = strMsg & "Há GR´s com status de ""inicial"" ou ""movimento após o fechamento"". Feche estas GR´s para conseguir fechar o turno." & vbCrLf
    blnSetarFocoControle = False
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[ValidaCamposFechaTurno]"
    ValidaCamposFechaTurno = False
  Else
    ValidaCamposFechaTurno = True
  End If
  Set objGeral = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.ValidaCamposFechaTurno]", _
            Err.Description
End Function

Public Sub FechamentoTurno()
  On Error GoTo trata
  Dim lngTURNOANTERIORID  As Long
  Dim strMsg              As String
  Dim objTurno            As busSisMed.clsTurno
  Dim strSql              As String
  Dim strData             As String
  '
  gsNomeUsuLib = ""
  If RetornaCodTurnoCorrente = 0 Then
    TratarErroPrevisto "Não há turno aberto", "frmUserTurnoInc.cmdFechar_Click"
    Exit Sub
  End If
  If Not ValidaCamposFechaTurno Then Exit Sub
  'Data1.Recordset!StatusOp = iStatusOp
  If MsgBox("Confirma fechamento do Turno corrente?", vbYesNo, TITULOSISTEMA) = vbNo Then Exit Sub
  'Encerrar Turno
  '
'''  If gsNomeUsuLib = "" Then
'''    If gbPedirSenhaFechaTurno = True Then
'''      '----------------------------
'''      '----------------------------
'''      'Pede Senha Superior (Diretor, Gerente ou Administrador
'''      If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
'''        'Só pede senha superior se quem estiver logado não for superior
'''        frmUserLoginSup.Show vbModal
'''
'''        If Len(Trim(gsNomeUsuLib)) = 0 Then
'''          strMsg = "Para efetuar o Fechamento/Abertura do Turno é necessário a Confirmação com senha superior."
'''          TratarErroPrevisto strMsg, "cmdConfirmar_Click"
'''          Exit Sub
'''        End If
'''        '
'''        'Capturou Nome do Usuário, continua processo de Sangria
'''      Else
'''        gsNomeUsuLib = gsNomeUsu
'''      End If
'''      '--------------------------------
'''      '--------------------------------
'''    End If
'''  End If
  'Tratamento dos botões
  DoEvents
  '
  '------------------------------------
  'FECHAMENTO
  '------------------------------------
  lngTURNOANTERIORID = RetornaCodTurnoCorrente
  
  Set objTurno = New busSisMed.clsTurno
  strData = DataHoraAtualFormatada
  'Fechar Turno corrente
  objTurno.FecharTurno lngTURNOANTERIORID, _
                       False, _
                       strData

  
  '
  IMP_COMPROV_FECHA_TURNO lngTURNOANTERIORID, gsNomeEmpresa, 1
  '
  MsgBox "O turno foi fechado com sucesso!", vbExclamation, TITULOSISTEMA
  '
  Set objTurno = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.FechamentoTurno]", _
            Err.Description
End Sub

  
'Propósito: Capturar as configurações do sistema
Public Sub Captura_Config()
  On Error GoTo trata
  Dim strSql            As String
  Dim objRs             As ADODB.Recordset
  Dim strDiaDaSemana    As String
  Dim objGeral          As busSisMed.clsGeral
  '
  Set objGeral = New busSisMed.clsGeral
  '
  strSql = "Select min(PKID) PKID from Configuracao "
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  glMaxSequencialImp = 9999
  If Not objRs.EOF Then
    lngCONFIGURACAOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  
  strSql = "Select * from Configuracao where PKID = " & Formata_Dados(lngCONFIGURACAOID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  If objRs.EOF Then
    gsPathLocal = ""
    gsPathLocalBackup = ""
    gsPathRede = ""
  
    giMaxDiasAtend = 0
    gbTrabImpA5 = False
'''    gbTrabComEstInterAutomatico = False
  Else
    '---------
    gsPathLocal = objRs.Fields("PathLocal").Value & ""
    gsPathLocalBackup = objRs.Fields("PathLocalBackup").Value & ""
    gsPathRede = objRs.Fields("PathRede").Value & ""

    If IsNumeric(objRs.Fields("QTDMAXIMADIASATEND").Value) Then
      giMaxDiasAtend = objRs.Fields("QTDMAXIMADIASATEND").Value
    End If
    gbTrabImpA5 = objRs.Fields("TRABALHACOMIMPRESSAOA5").Value
    
'''
'''    If Not IsNull(objRs.Fields("DIAFECHAFOLHA").Value) Then
'''      glDiaFechaFolha = objRs.Fields("DIAFECHAFOLHA").Value
'''    Else
'''      glDiaFechaFolha = 1
'''    End If
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set objGeral = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.Captura_Config]", _
            Err.Description
End Sub


Public Sub Retorna_totais_GR(ByVal lngGRID As Long, _
                             ByRef curVrPrest As Currency, _
                             ByRef curVrCasa As Currency, _
                             ByRef curVrTotal As Currency)
  On Error GoTo trata
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim objGer    As busSisMed.clsGeral
  '
  'Tratar campos
  Set objGer = New busSisMed.clsGeral
  '
  strSql = "SELECT  sum(vw_cons_t_Financ.PgtoTotal) as PgtoTotal, " & _
      " sum(vw_cons_t_Financ.FINALCASACONSESPECIE) as FINALCASACONSESPECIE, sum(vw_cons_t_Financ.FINALPRESTCONSESPECIE) as FINALPRESTCONSESPECIE, sum(vw_cons_t_Financ.FINALDONORXCONSESPECIE) as FINALDONORXCONSESPECIE, sum(vw_cons_t_Financ.FINALTECRXCONSESPECIE) as FINALTECRXCONSESPECIE, " & _
      " sum(vw_cons_t_Financ.FINALDONOULTRACONSESPECIE) as FINALDONOULTRACONSESPECIE, sum(vw_cons_t_Financ.FINALCASACONSCARTAO) as FINALCASACONSCARTAO, sum(vw_cons_t_Financ.FINALPRESTCONSCARTAO) as FINALPRESTCONSCARTAO, " & _
      " sum(vw_cons_t_Financ.FINALDONORXCONSCARTAO) as FINALDONORXCONSCARTAO, sum(vw_cons_t_Financ.FINALTECRXCONSCARTAO) as FINALTECRXCONSCARTAO, sum(vw_cons_t_Financ.FINALDONOULTRACONSCARTAO) as FINALDONOULTRACONSCARTAO, " & _
      " sum(vw_cons_t_Financ.FINALCASACONSCONVENIO) as FINALCASACONSCONVENIO, sum(vw_cons_t_Financ.FINALPRESTCONSCONVENIO) as FINALPRESTCONSCONVENIO, " & _
      " sum(vw_cons_t_Financ.FINALDONORXCONSCONVENIO) as FINALDONORXCONSCONVENIO, sum(vw_cons_t_Financ.FINALTECRXCONSCONVENIO) as FINALTECRXCONSCONVENIO, sum(vw_cons_t_Financ.FINALDONOULTRACONSCONVENIO) as FINALDONOULTRACONSCONVENIO, " & _
      " sum(vw_cons_t_Financ.FINALCASAESPECIE) as FINALCASAESPECIE, sum(vw_cons_t_Financ.FINALCASACARTAO) as FINALCASACARTAO, sum(vw_cons_t_Financ.FINALCASACONVENIO) as FINALCASACONVENIO, sum(vw_cons_t_Financ.FINALPRESTESPECIE) as FINALPRESTESPECIE, sum(vw_cons_t_Financ.FINALPRESTCONVENIO) as FINALPRESTCONVENIO, sum(vw_cons_t_Financ.FINALPRESTCARTAONAOACEITA) as FINALPRESTCARTAONAOACEITA, " & _
      " sum(vw_cons_t_Financ.FINALPRESTCARTAOACEITAPGFUTURO) as FINALPRESTCARTAOACEITAPGFUTURO, sum(vw_cons_t_Financ.FINALDONORXESPECIE) as FINALDONORXESPECIE, sum(vw_cons_t_Financ.FINALDONORXCONVENIO) as FINALDONORXCONVENIO, sum(vw_cons_t_Financ.FINALDONORXCARTAONAOACEITA) as FINALDONORXCARTAONAOACEITA, " & _
      " sum(vw_cons_t_Financ.FINALDONORXCARTAOACEITAPGFUTURO) as FINALDONORXCARTAOACEITAPGFUTURO, sum(vw_cons_t_Financ.FINALDONOULTRAESPECIE) as FINALDONOULTRAESPECIE, sum(vw_cons_t_Financ.FINALDONOULTRACONVENIO) as FINALDONOULTRACONVENIO, sum(vw_cons_t_Financ.FINALDONOULTRACARTAONAOACEITA) as FINALDONOULTRACARTAONAOACEITA, " & _
      " sum(vw_cons_t_Financ.FINALDONOULTRACARTAOACEITAPGFUTURO) as FINALDONOULTRACARTAOACEITAPGFUTURO, sum(vw_cons_t_Financ.FINALTECRXESPECIE) as FINALTECRXESPECIE, sum(vw_cons_t_Financ.FINALTECRXCONVENIO) as FINALTECRXCONVENIO, sum(vw_cons_t_Financ.FINALTECRXCARTAONAOACEITA) as FINALTECRXCARTAONAOACEITA, " & _
      " sum(vw_cons_t_Financ.FINALTECRXCARTAOACEITAPGFUTURO) as FINALTECRXCARTAOACEITAPGFUTURO " & _
      " FROM GR " & _
      " INNER JOIN vw_cons_t_Financ ON GR.PKID = vw_cons_t_Financ.GRID " & _
      " WHERE GR.PKID = " & Formata_Dados(lngGRID, tpDados_Longo) & _
      " GROUP BY GR.PKID"

  'strSql = "SELECT  sum(vw_cons_t_Financ.PgtoTotal) as PgtoTotal, " & _
      " sum(vw_cons_t_Financ.FINALCASACONSESPECIE) as FINALCASACONSESPECIE, sum(vw_cons_t_Financ.FINALPRESTCONSESPECIE) as FINALPRESTCONSESPECIE, sum(vw_cons_t_Financ.FINALDONORXCONSESPECIE) as FINALDONORXCONSESPECIE, sum(vw_cons_t_Financ.FINALTECRXCONSESPECIE) as FINALTECRXCONSESPECIE, " & _
      " sum(vw_cons_t_Financ.FINALDONOULTRACONSESPECIE) as FINALDONOULTRACONSESPECIE, sum(vw_cons_t_Financ.FINALCASACONSCARTAO) as FINALCASACONSCARTAO, sum(vw_cons_t_Financ.FINALPRESTCONSCARTAO) as FINALPRESTCONSCARTAO, " & _
      " sum(vw_cons_t_Financ.FINALDONORXCONSCARTAO) as FINALDONORXCONSCARTAO, sum(vw_cons_t_Financ.FINALTECRXCONSCARTAO) as FINALTECRXCONSCARTAO, sum(vw_cons_t_Financ.FINALDONOULTRACONSCARTAO) as FINALDONOULTRACONSCARTAO, " & _
      " sum(vw_cons_t_Financ.FINALCASACONSCONVENIO) as FINALCASACONSCONVENIO, sum(vw_cons_t_Financ.FINALPRESTCONSCONVENIO) as FINALPRESTCONSCONVENIO, " & _
      " sum(vw_cons_t_Financ.FINALDONORXCONSCONVENIO) as FINALDONORXCONSCONVENIO, sum(vw_cons_t_Financ.FINALTECRXCONSCONVENIO) as FINALTECRXCONSCONVENIO, sum(vw_cons_t_Financ.FINALDONOULTRACONSCONVENIO) as FINALDONOULTRACONSCONVENIO, " & _
      " sum(vw_cons_t_Financ.FINALCASAESPECIE) as FINALCASAESPECIE, sum(vw_cons_t_Financ.FINALCASACARTAO) as FINALCASACARTAO, sum(vw_cons_t_Financ.FINALCASACONVENIO) as FINALCASACONVENIO, sum(vw_cons_t_Financ.FINALPRESTESPECIE) as FINALPRESTESPECIE, sum(vw_cons_t_Financ.FINALPRESTCONVENIO) as FINALPRESTCONVENIO, sum(vw_cons_t_Financ.FINALPRESTCARTAONAOACEITA) as FINALPRESTCARTAONAOACEITA, " & _
      " sum(vw_cons_t_Financ.FINALPRESTCARTAOACEITAPGFUTURO) as FINALPRESTCARTAOACEITAPGFUTURO, sum(vw_cons_t_Financ.FINALDONORXESPECIE) as FINALDONORXESPECIE, sum(vw_cons_t_Financ.FINALDONORXCONVENIO) as FINALDONORXCONVENIO, sum(vw_cons_t_Financ.FINALDONORXCARTAONAOACEITA) as FINALDONORXCARTAONAOACEITA, " & _
      " sum(vw_cons_t_Financ.FINALDONORXCARTAOACEITAPGFUTURO) as FINALDONORXCARTAOACEITAPGFUTURO, sum(vw_cons_t_Financ.FINALDONOULTRAESPECIE) as FINALDONOULTRAESPECIE, sum(vw_cons_t_Financ.FINALDONOULTRACONVENIO) as FINALDONOULTRACONVENIO, sum(vw_cons_t_Financ.FINALDONOULTRACARTAONAOACEITA) as FINALDONOULTRACARTAONAOACEITA, " & _
      " sum(vw_cons_t_Financ.FINALDONOULTRACARTAOACEITAPGFUTURO) as FINALDONOULTRACARTAOACEITAPGFUTURO, sum(vw_cons_t_Financ.FINALTECRXESPECIE) as FINALTECRXESPECIE, sum(vw_cons_t_Financ.FINALTECRXCONVENIO) as FINALTECRXCONVENIO, sum(vw_cons_t_Financ.FINALTECRXCARTAONAOACEITA) as FINALTECRXCARTAONAOACEITA, " & _
      " sum(vw_cons_t_Financ.FINALTECRXCARTAOACEITAPGFUTURO) as FINALTECRXCARTAOACEITAPGFUTURO " & _
      " FROM GRPAGAMENTO INNER JOIN GRPGTO ON GRPAGAMENTO.PKID = GRPGTO.GRPAGAMENTOID " & _
      " INNER JOIN GR ON GR.PKID = GRPGTO.GRID " & _
      " INNER JOIN vw_cons_t_Financ ON GR.PKID = vw_cons_t_Financ.GRID " & _
      " WHERE GRPAGAMENTO.PKID = " & Formata_Dados(lngGRPAGAMENTOID, tpDados_Longo) & _
      " GROUP BY GRPAGAMENTO.PKID"

  '
  curVrPrest = 0
  curVrCasa = 0
  curVrTotal = 0
  '
  Set objRs = objGer.ExecutarSQL(strSql)
  '
  If Not objRs.EOF Then   'se já houver algum item
    curVrCasa = objRs.Fields("FINALCASACONSESPECIE").Value + _
                objRs.Fields("FINALCASACONSCARTAO").Value + _
                objRs.Fields("FINALCASACONSCONVENIO").Value + _
                objRs.Fields("FINALCASAESPECIE").Value + _
                objRs.Fields("FINALCASACARTAO").Value + _
                objRs.Fields("FINALCASACONVENIO").Value
    '
    curVrTotal = objRs.Fields("PgtoTotal").Value
    '
    'TOTAIS Prestador
    curVrPrest = objRs.Fields("FINALPRESTCONSESPECIE").Value + _
                 objRs.Fields("FINALPRESTCONSCARTAO").Value + _
                 objRs.Fields("FINALPRESTCONSCONVENIO").Value + _
                 objRs.Fields("FINALPRESTESPECIE").Value + _
                 objRs.Fields("FINALPRESTCONVENIO").Value + _
                 objRs.Fields("FINALPRESTCARTAONAOACEITA").Value
    'TOTAIS Dono RX
    curVrPrest = curVrPrest + _
                 objRs.Fields("FINALDONORXCONSESPECIE").Value + _
                 objRs.Fields("FINALDONORXCONSCARTAO").Value + _
                 objRs.Fields("FINALDONORXCONSCONVENIO").Value + _
                 objRs.Fields("FINALDONORXESPECIE").Value + _
                 objRs.Fields("FINALDONORXCONVENIO").Value + _
                 objRs.Fields("FINALDONORXCARTAONAOACEITA").Value
    'TOTAIS Tecnico RX
    curVrPrest = curVrPrest + _
                 objRs.Fields("FINALTECRXCONSESPECIE").Value + _
                 objRs.Fields("FINALTECRXCONSCARTAO").Value + _
                 objRs.Fields("FINALTECRXCONSCONVENIO").Value + _
                 objRs.Fields("FINALTECRXESPECIE").Value + _
                 objRs.Fields("FINALTECRXCONVENIO").Value + _
                 objRs.Fields("FINALTECRXCARTAONAOACEITA").Value
    'TOTAIS Dono Ultra
    curVrPrest = curVrPrest + _
                 objRs.Fields("FINALDONOULTRACONSESPECIE").Value + _
                 objRs.Fields("FINALDONOULTRACONSCARTAO").Value + _
                 objRs.Fields("FINALDONOULTRACONSCONVENIO").Value + _
                 objRs.Fields("FINALDONOULTRAESPECIE").Value + _
                 objRs.Fields("FINALDONOULTRACONVENIO").Value + _
                 objRs.Fields("FINALDONOULTRACARTAONAOACEITA").Value
    'TOTAIS A RECEBER
    curVrPrest = curVrPrest + _
                 objRs.Fields("FINALPRESTCARTAOACEITAPGFUTURO").Value + _
                 objRs.Fields("FINALDONORXCARTAOACEITAPGFUTURO").Value + _
                 objRs.Fields("FINALTECRXCARTAOACEITAPGFUTURO").Value + _
                 objRs.Fields("FINALDONOULTRACARTAOACEITAPGFUTURO").Value
    '
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGer = Nothing
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Public Sub TratarImagemScanner(datData As Date, _
                                strPathOrigem As String, _
                                strPathDestinoBkp As String, _
                                strArquivoOrigem As String, _
                                lngPRONTUARIOID As Long, _
                                lngGRID As Long, _
                                ByRef strPathDestino As String, _
                                ByRef strArquivoDestino As String)
  On Error GoTo trata
  '
  Dim objFso As FileSystemObject
  'TRATAR DIRETÓRIO FINAL PARA COPIAR ARQUIVO
  '
  Set objFso = New FileSystemObject
  strPathDestino = gsPathRede & lngPRONTUARIOID & "\"
  If Not objFso.FolderExists(strPathDestino) Then
    'Diretório não existe, cria diretório
    objFso.CreateFolder (strPathDestino)
  End If
  'Verifica extensão do arquivo
  If lngGRID <> 0 Then
    strArquivoDestino = lngGRID & "_" & Format(datData, "YYYYMMDDhhmmss") & "." & RetornaEtensaoArquivo(strArquivoOrigem)
  Else
    strArquivoDestino = strArquivoDestino
  End If
  '
  'COPIA ARQUIVO
  objFso.CopyFile strPathOrigem & strArquivoOrigem, strPathDestino & strArquivoDestino
  'COPIA DESEGURANÇA
  objFso.CopyFile strPathOrigem & strArquivoOrigem, strPathDestinoBkp & strArquivoDestino
  objFso.DeleteFile strPathOrigem & strArquivoOrigem
  
  Set objFso = Nothing
  
'''    gsPathLocal = objRs.Fields("PathLocal").Value & ""
'''    gsPathLocalBackup = objRs.Fields("PathLocalBackup").Value & ""
'''    gsPathRede = objRs.Fields("PathRede").Value & ""
  
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function RetornaEtensaoArquivo(strArquivoDestino As String)
  On Error GoTo trata
  '
  Dim strExtensao As String
  strExtensao = Mid(strArquivoDestino, InStrRev(strArquivoDestino, ".") + 1)
  '
  RetornaEtensaoArquivo = strExtensao
  Exit Function
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Function


Public Sub Importar_Receitas(strPathOrigem As String)
  On Error GoTo trata
  Dim objFso          As Scripting.FileSystemObject
  Dim objFolder       As Scripting.Folder
  Dim objFile         As Scripting.File
  Dim strFileName     As String
  
  Dim datData         As Date
  Dim strData         As String
  Dim strErroGeral    As String
  Dim strSucessoGeral As String
  Dim strProntuarioId As String
  Dim strDataAtend    As String
  Dim lngPRONTUARIOID As Long
  Dim lngATENDIMENTOID As Long
  '
  Dim objAtendimento    As busSisMed.clsAtendimento
  Dim objGR             As busSisMed.clsGR
  Dim gsPathFinal       As String
  '
  'TRATAR DIRETÓRIO FINAL PARA COPIAR ARQUIVO
  '
  datData = Now
  strData = Format(datData, "DD/MM/YYYY hh:mm")

  Set objFso = New FileSystemObject
  '
  Set objFolder = objFso.GetFolder(strPathOrigem)
  strErroGeral = ""
  strSucessoGeral = ""
  For Each objFile In objFolder.Files
    'Para cada arquivo de imagem no diretório
    strFileName = objFile.Name
    If ValidaCamposArquivoScanner(strFileName, _
                                  strProntuarioId, _
                                  strDataAtend, _
                                  strErroGeral, _
                                  lngPRONTUARIOID, _
                                  lngATENDIMENTOID) Then
      
      Set objAtendimento = New busSisMed.clsAtendimento
      Set objGR = New busSisMed.clsGR
      
      TratarImagemScanner datData, _
                          gsPathLocal, _
                          gsPathLocalBackup, _
                          strFileName, _
                          lngPRONTUARIOID, _
                          0, _
                          gsPathFinal, _
                          strFileName
      
      If lngATENDIMENTOID <> 0 Then
        'Alterar Atendimento
        objAtendimento.AlterarAtendimento lngATENDIMENTOID, _
                                          gsPathFinal, _
                                          strFileName, _
                                          ""
        '
      ElseIf lngATENDIMENTOID = 0 Then
        'Inserir Atendimento
        objAtendimento.InserirAtendimento 0, _
                                          strDataAtend, _
                                          "A", _
                                          gsPathFinal, _
                                          strFileName, _
                                          "", _
                                          strData, _
                                          lngPRONTUARIOID
      End If
      Set objAtendimento = Nothing
      Set objGR = Nothing
      '
      If strSucessoGeral = "" Then
        strSucessoGeral = "As seguintes receitas scanneadas foram importadas com sucesso:" & vbCrLf & vbCrLf
      End If
      strSucessoGeral = strSucessoGeral & strFileName & vbCrLf
    End If
  Next
  
  If strSucessoGeral <> "" Then
    MsgBox strSucessoGeral, vbExclamation, TITULOSISTEMA
  End If
  If strErroGeral <> "" Then
    TratarErroPrevisto strErroGeral, "[mdlUserFuncion.ValidaCamposArquivoScanner]"
  End If
  
  
  
  Set objFolder = Nothing
  Set objFso = Nothing
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCamposArquivoScanner(ByVal strFileName As String, _
                                            ByRef strProntuarioId As String, _
                                            ByRef strDataAtend As String, _
                                            ByRef strErro As String, _
                                            ByRef lngPRONTUARIOID As Long, _
                                            ByRef lngATENDIMENTOID As Long) As Boolean
  On Error GoTo trata
  Dim ValidaCampos    As Boolean
  Dim strMsg          As String
  Dim strEtensaoArquivo  As String
  Dim strDataAtendFinal   As String
  '
  Dim objGeral                  As busSisMed.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  '
  ValidaCampos = False
  'Obrigatório apenas para arquivista
  If strFileName = "" Then
    strMsg = strMsg & vbTab & vbTab & "Nome do arquivo não informado" & vbCrLf
  End If
  If InStr(strFileName, "_") = 0 Then
    strMsg = strMsg & vbTab & vbTab & "Separador [_] não definido" & vbCrLf
  End If
  'Obtem os dados separados
  strEtensaoArquivo = RetornaEtensaoArquivo(strFileName)
  
  If InStr(strFileName, "_") <> 0 Then
    strProntuarioId = Left(strFileName, InStr(strFileName, "_") - 1)
    strDataAtend = Mid(strFileName, InStr(strFileName, "_") + 1, Len(strFileName) - (InStr(strFileName, "_") + 1) - (Len(strEtensaoArquivo)))
  Else
    strProntuarioId = ""
    strDataAtend = ""
  End If
  If Not IsNumeric(strProntuarioId) Then
    strMsg = strMsg & vbTab & vbTab & "Número do prontuário [" & strProntuarioId & "] não numérico" & vbCrLf
  End If
  
  If Len(strDataAtend) <> 12 Then
    strMsg = strMsg & vbTab & vbTab & "Data de atendimento [" & strDataAtend & "] tem que possuir 12 caracteres no formato DDMMYYYYhhmm" & vbCrLf
  End If
  strDataAtendFinal = Left(strDataAtend, 2) & "/" & _
        Mid(strDataAtend, 3, 2) & "/" & _
        Mid(strDataAtend, 5, 4) & " " & _
        Mid(strDataAtend, 9, 2) & ":" & _
        Mid(strDataAtend, 11, 2)
  If Not Valida_Data_Str(strDataAtendFinal, TpObrigatorio) Then
    strMsg = strMsg & vbTab & vbTab & "Data de atendimento [" & strDataAtendFinal & "] não é uma data válida" & vbCrLf
  End If
  strDataAtend = strDataAtendFinal
  Set objGeral = New busSisMed.clsGeral
  If Len(strMsg) = 0 Then
    'Validação na base de dados
    'PRONTUARIO
    lngPRONTUARIOID = 0
    strSql = "SELECT PKID FROM PRONTUARIO WHERE PRONTUARIO.PKID = " & Formata_Dados(strProntuarioId, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngPRONTUARIOID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  If lngPRONTUARIOID = 0 Then
    strMsg = strMsg & vbTab & vbTab & "Número do prontuário [" & strProntuarioId & "] não cadastrado na base de dados" & vbCrLf
  End If
  If Len(strMsg) = 0 Then
    'Verifica se já houve atendimento apra aquele prontuário
    'ATENDIMENTO
    lngATENDIMENTOID = 0
    strSql = "SELECT PKID FROM ATENDIMENTO WHERE ATENDIMENTO.PRONTUARIOID = " & Formata_Dados(lngPRONTUARIOID, tpDados_Longo) & _
          " AND DATA = " & Formata_Dados(strDataAtend, tpDados_DataHora)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngATENDIMENTOID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  Set objGeral = Nothing
  If Len(strMsg) <> 0 Then
    If strErro <> "" Then
      strErro = strErro & vbCrLf & vbCrLf
    Else
      strErro = "Foram encontrados erros na composição das seguintes receitas scanneadas: " & vbCrLf & vbCrLf
    End If
    strErro = strErro & strFileName & vbCrLf
    strErro = strErro & strMsg
    'TratarErroPrevisto strMsg, "[mdlUserFuncion.ValidaCamposArquivoScanner]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  ValidaCamposArquivoScanner = ValidaCampos
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[mdlUserFuncion.ValidaCamposArquivoScanner]", _
            Err.Description


End Function

