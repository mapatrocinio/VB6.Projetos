Attribute VB_Name = "mdlUserFunction"
Option Explicit

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

'preencher combo box
Public Sub PreencheCombo(objCbo, _
                         ByVal strSql As String, _
                         Optional TipoTodos As Boolean = True, _
                         Optional TipoBranco As Boolean = False, _
                         Optional SelPriItem As Boolean = False)
  On Error GoTo trata
  Dim objRs         As ADODB.Recordset
  Dim objGeral      As busApler.clsGeral
  Dim blnPrimeiro   As Boolean
  Dim strPriItem    As String
  '
  Set objGeral = New busApler.clsGeral
  '
  blnPrimeiro = True
  Set objRs = objGeral.ExecutarSQL(strSql)
  objCbo.Clear
  If TipoBranco Then _
     objCbo.AddItem ""
  If TipoTodos Then _
     objCbo.AddItem "<TODOS>"
  Do While Not objRs.EOF
    objCbo.AddItem objRs.Fields(0) & ""
    If blnPrimeiro Then strPriItem = objRs.Fields(0) & ""
    blnPrimeiro = False
    objRs.MoveNext
  Loop
  If SelPriItem And strPriItem <> "" Then objCbo.Text = strPriItem
  Set objGeral = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.PreencheCombo]", _
            Err.Description
End Sub


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

Public Sub TratarErro(ByVal pNumero As Long, _
                    ByVal pDescricao As String, _
                    ByVal pSource As String)
  '
  On Error Resume Next
  Dim strUsuario As String
  Dim intI       As Integer
    
  intI = FreeFile
  Open App.Path & "\Apler.txt" For Append As #intI
  
  Print #intI, Format(Now(), "DD/MM/YYYY hh:mm") & ";" & pNumero & ";" & pSource & ";" & pDescricao
  Close #intI
  'mostrar Mensagem
  MsgBox "O Seginte Erro Ocorreu: " & vbCrLf & vbCrLf & _
    "Número: " & pNumero & vbCrLf & _
    "Descrição: " & pDescricao & vbCrLf & vbCrLf & _
    "Origem: " & pSource & vbCrLf & _
    "Data/Hora: " & Format(Now(), "DD/MM/YYYY hh:mm") & vbCrLf & _
    "Erro gravado no arquivo: " & App.Path & "\Apler.txt" & vbCrLf & vbCrLf & _
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
  frmMDI.snuRelatorio(0).Visible = False
  '
  frmMDI.mnuArquivo(2).Visible = False
  'frmMDI.mnuArquivo(3).Visible = False
  frmMDI.mnuArquivo(4).Visible = False
  frmMDI.mnuArquivo(5).Visible = False
  frmMDI.mnuCadastro.Visible = False
    
  '
  If pAcao = 1 Then 'Monta Conexão
    Select Case Trim(gsNivel)
    Case gsAdmin
      'Acesso Completo
      frmMDI.mnuCadastro.Visible = True
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

