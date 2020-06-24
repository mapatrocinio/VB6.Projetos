<%
function RetornaDataPerAuditoria(dtaInicio, dtaTermino)
	dim strDataInicio
	dim strDataTermino
	strDataInicio = day(dtaInicio)
	if month(dtaInicio) <> month(dtaTermino) or _
		year(dtaInicio) <> year(dtaTermino) then
		strDataInicio = strDataInicio & " de " & RetornaMesExtenso(month(dtaInicio))
	end if		
	if year(dtaInicio) <> year(dtaTermino) then
		strDataInicio = strDataInicio & " de " & year(dtaInicio)
	end if		
	strDataTermino = day(strDataTermino) & " de " & _
		RetornaMesExtenso(month(dtaTermino)) & _
		" de " & year(dtaTermino)
	RetornaDataPerAuditoria = strDataInicio & " a " & strDataTermino
	
end function
function RetornaMesExtenso(pMes)
	Select case pMes
	case 1: RetornaMesExtenso = "Janeiro"
	case 2: RetornaMesExtenso = "Fevereiro"
	case 3: RetornaMesExtenso = "Março"
	case 4: RetornaMesExtenso = "Abril"
	case 5: RetornaMesExtenso = "Maio"
	case 6: RetornaMesExtenso = "Junho"
	case 7: RetornaMesExtenso = "Julho"
	case 8: RetornaMesExtenso = "Agosto"
	case 9: RetornaMesExtenso = "Setembro"
	case 10: RetornaMesExtenso = "Outubro"
	case 11: RetornaMesExtenso = "Novembro"
	case 12: RetornaMesExtenso = "Dezembro"
	end select
end function

Sub TratarIdentacao(blnIdentar)
	if blnIdentar = true then
	%>
    With objWord.Selection.ParagraphFormat
        .LeftIndent = 56.69291
        .RightIndent = 0
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = 0
        .Alignment = 3
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = -56.69291
        .OutlineLevel = 10
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
<%else%>
    With objWord.Selection.ParagraphFormat
        .LeftIndent = 0
        .RightIndent = 0
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = 0
        .Alignment = 3
        .WidowControl = True
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = 0
        .OutlineLevel = 10
        .CharacterUnitLeftIndent = 0
        .CharacterUnitRightIndent = 0
        .CharacterUnitFirstLineIndent = 0
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With

<%end if
End sub
Public Function RetornaQtdOcorrencias(strTexto, strStringFind)
  Dim intRetorno
  Dim blnAchou
  Dim intPosicao
  intRetorno = 0
  intPosicao = InStr(1, strTexto, strStringFind)
  If intPosicao > 0 Then
    blnAchou = True
  Else
    blnAchou = False
  End If
  Do While blnAchou
    intRetorno = intRetorno + 1
    intPosicao = InStr(intPosicao + 1, strTexto, strStringFind)
    If intPosicao > 0 Then
      blnAchou = True
    Else
      blnAchou = False
    End If
  Loop
  RetornaQtdOcorrencias = intRetorno
End function
Sub exibir_planilha_item(ByVal objConn, _
                         ByVal seq_relatorio, _
												 ByVal status_detalhes, _
												 ByVal ind_tabela)

  Dim qtdLinhasTabela
  Dim qtdColunasTabela
  Dim strTabela
  Dim intCol
  Dim intLin

  Dim objRs
  Dim objRs1
  Dim objRs2
  Dim objRs3
  Dim objRsPrincipal

  Dim seq_planilha
  Dim descr_titulo
  Dim descr_nivel_interno
  Dim descr_nivel_mediano
  Dim descr_nivel_externo
  Dim ind_totalizador
  Dim seq_coluna
  Dim vetCol()
  Dim vetLin()
  Dim strValor
  Dim strValorTotalizador
  Set objRsPrincipal = CreateObject("ADODB.Recordset")
	if a <> "" then
		response.write a
		response.end
	end if
  'If seq_relatorio <> 4 Then Exit Sub
  Select Case ind_tabela
  Case "template_relatorio"
    'Tabela de template
    strSql = "select planilha.* from planilha " & _
        "inner join planilha_template_relatorio " & _
        " on planilha.seq_planilha = planilha_template_relatorio.seq_planilha " & _
        "where planilha_template_relatorio.seq_planilha = " & seq_relatorio & _
        " order by planilha_template_relatorio.num_ordem"
  Case "relatorio_sa"
    'Tabela de relatorio_auditoria
    strSql = "select planilha.* from planilha " & _
        "inner join planilha_sa_item_auditoria " & _
        " on planilha.seq_planilha = planilha_sa_item_auditoria.seq_planilha " & _
        "where ano_processo = " & ano_processo & _
        " and sig_orgao_processo = '" & sig_orgao_processo & "'" & _
        " and seq_processo = " & seq_processo & _
        " and seq_sa = " & seq_sa & _
				" and seq_sa_complementar = '" & seq_sa_complementar & "'" & _
        " and seq_assunto = " & seq_assunto & _
				" and seq_item_sa = " & seq_item_sa & _
				" and seq_area = " & seq_area & _
        " order by planilha_sa_item_auditoria.num_ordem"
  Case "relatorio_salvar_listar"
    'Tabela de relatorio_auditoria
    strSql = "select planilha.* from planilha " & _
        "inner join planilha_relatorio_auditoria " & _
        " on planilha.seq_planilha = planilha_relatorio_auditoria.seq_planilha " & _
        "where ano_processo = " & ano_processo & _
        " and sig_orgao_processo = '" & sig_orgao_processo & "'" & _
        " and seq_processo = " & seq_processo & _
        " and seq_relatorio = " & seq_relatorio & _
        " order by planilha_relatorio_auditoria.num_ordem"
  Case "relatorio_PP"
    'Tabela de relatorio_PP
    strSql = "select planilha.* from planilha " & _
        "inner join planilha_comentario_item_sa " & _
        " on planilha.seq_planilha = planilha_comentario_item_sa.seq_planilha " & _
        "where ano_processo = " & ano_processo & _
        " and sig_orgao_processo = '" & sig_orgao_processo & "'" & _
        " and seq_processo = " & seq_processo & _
        " and seq_sa = " & seq_sa & _
				" and seq_sa_complementar = '" & seq_sa_complementar & "'" & _
        " and seq_assunto = " & seq_assunto & _
				" and seq_item_sa = " & seq_item_sa & _
				" and seq_area = " & seq_area & _
				" and seq_comentario = " & seq_comentario & _
        " order by planilha_comentario_item_sa.num_ordem"
  End Select
  objRsPrincipal.Open strSql, objConn, 0, 1
    
  If status_detalhes = "inicial" Then
    'Se estiver no estado inicial da página,
    'pegar os dados do banco de dados, se houver

    strTabela = ""
    If Not objRsPrincipal.EOF Then
			blnPossuiPlanilha = true
      Do While Not objRsPrincipal.EOF
          seq_planilha = objRsPrincipal.Fields("seq_planilha")
          Set objRs = CreateObject("ADODB.Recordset")
          strSql = "select * from planilha where seq_planilha = " & seq_planilha
          objRs.Open strSql, objConn, 0, 1
          If Not objRs.EOF Then
              'Há planilha cadastrada para este
              'Em um primeiro momento preenche os dados
              '
              descr_titulo = objRs.Fields("descr_titulo") & ""
              qtdLinhasTabela = objRs.Fields("qtd_linha") + objRs.Fields("qtd_nivel_coluna") '+ 1
              qtdColunasTabela = objRs.Fields("qtd_coluna") + objRs.Fields("qtd_nivel_linha") '+ 1
              ind_tipo_planilha = objRs.Fields("ind_tipo_planilha") & ""
							'Formatação
							ind_exibir_titulo	= objRs.Fields("ind_exibir_titulo") & ""
							ind_alin_tit_lin	= objRs.Fields("ind_alinhamento_titulo_linha") & ""
							ind_alin_tit_col	= objRs.Fields("ind_alinhamento_titulo_coluna") & ""
							ind_alin_dados		= objRs.Fields("ind_alinhamento_dados") & ""
							ind_neg_tit_lin		= objRs.Fields("ind_negrito_titulo_linha") & ""
							ind_neg_tit_col		= objRs.Fields("ind_negrito_titulo_coluna") & ""
							ind_neg_dados			= objRs.Fields("ind_negrito_dados") & ""
              'Definidos as linhas/colunas da tabela monta a tabela
              'em variável
							%>
							objWord.Selection.MoveRight <%=wdCharacter%>, 1
							objWord.Selection.TypeParagraph
							
							if "<%=ind_exibir_titulo%>" = "S" then
								
								objWord.Selection.ParagraphFormat.Alignment = <%=wdAlignParagraphCenter%>
								objWord.Selection.Font.Name = "Arial"
								objWord.Selection.Font.Size = 9
								objWord.Selection.TypeText "<%=descr_titulo%>"
								
								objWord.Selection.MoveRight <%=wdCharacter%>, 1
								objWord.Selection.TypeParagraph
							end if                    
							
              objWordDoc.Tables.Add objWord.Selection.Range, 1, 1 ', <%=wdWord9TableBehavior%>,<%=wdAutoFitContent%>
							'objWord.Selection.Tables(1).AutoFitBehavior (<%=wdAutoFitContent%>)
							objWord.Selection.Tables(1).Rows.Alignment = <%=wdAlignRowCenter%>
              objWord.Selection.Tables(1).Columns(1).SetWidth 60, <%=wdAdjustProportional%>
              'Seleciona Coluna
              objWord.Selection.Tables(1).Columns(1).Select
              
              'On Error Resume Next
              <%      
              'De posse da tabela adiciona as linhas/colunas
              'com os valores provenientes da base de dados
              'através de replace nos tags previamente cadastrados
              Set objRs1 = CreateObject("ADODB.Recordset")
              Set objRs2 = CreateObject("ADODB.Recordset")
              Set objRs3 = CreateObject("ADODB.Recordset")
              'Abre recordsets
              'linhas
              strSql = "select * from linha where seq_planilha = " & seq_planilha
              objRs1.Open strSql, objConn, 0, 1
              'colunas
              strSql = "select * from coluna where seq_planilha = " & seq_planilha
              objRs2.Open strSql, objConn, 0, 1
              'linhas x colunas (valor_planilha)
              strSql = "select * from valor_planilha where seq_planilha = " & seq_planilha
              objRs3.Open strSql, objConn, 0, 1
              '
              For intLin = 1 To qtdLinhasTabela
                  'Abre linha da tabela
                  If intLin <> 1 Then
											%>
                      objWord.Selection.Tables(1).Rows.Add
                      <%
                  End If
                  For intCol = 1 To qtdColunasTabela
                      If intCol <> 1 And intLin = 1 Then
													%>
                          objWord.Selection.Tables(1).Columns.Add
                          <%
                      End If
                      %>
                      objWord.Selection.Tables(1).Cell(<%=intLin%>, <%=intCol%>).Select
                      <%
											'STOP
                      If (intCol <= objRs.Fields("qtd_nivel_linha") And intLin <= objRs.Fields("qtd_nivel_coluna")) Then 'Or _
                              '(intCol = qtdColunasTabela And intLin = qtdLinhasTabela) Then
                          'interseção entre linha x coluna do cabeçalho
                          'da tabela não adiciona nada
                          'strTabela = "" 'replace(strTabela, "<*" & intLin & "_" & intCol & "*>", "")
                      ElseIf intLin <= objRs.Fields("qtd_nivel_coluna") Then
                          'Cabeçalho da coluna
                          strValor = ""
                          strValorTotalizador = ""
                          If Not objRs2.EOF Then
                              objRs2.Filter = "seq_coluna = " & intCol - objRs.Fields("qtd_nivel_linha")
                              If Not objRs2.EOF Then
                                  If (intLin = 1 And objRs.Fields("qtd_nivel_coluna") = 1) Or _
                                      (intLin = 2 And objRs.Fields("qtd_nivel_coluna") = 2) Or _
                                      (intLin = 3 And objRs.Fields("qtd_nivel_coluna") = 3) Then
                                      strValor = objRs2.Fields("descr_nivel_interno") & ""
                                  ElseIf (intLin = 2 And objRs.Fields("qtd_nivel_coluna") = 3) Then
                                      strValor = objRs2.Fields("descr_nivel_mediano") & ""
                                  ElseIf (intLin = 1 And objRs.Fields("qtd_nivel_coluna") = 2) Or _
                                      (intLin = 1 And objRs.Fields("qtd_nivel_coluna") = 3) Or _
                                      (intLin = 2 And objRs.Fields("qtd_nivel_coluna") = 2) Or _
                                      (intLin = 3) Then
                                      strValor = objRs2.Fields("descr_nivel_externo") & ""
                                  End If
                                  If objRs2.Fields("ind_totalizador") & "" = "S" Then
                                      strValorTotalizador = " checked "
                                  End If
                              End If
                              objRs2.Filter = ""
                          End If
                          'strTabela = replace(strTabela, "<*" & intLin & "_" & intCol & "*>", strValor)
                          'strTabela = replace(strTabela, "|*" & qtdLinhasTabela & "_" & intCol & "*|", strValorTotalizador)
                          %>
                          objWord.Selection.Shading.Texture = <%=wdTexture10Percent%>
                          objWord.Selection.ParagraphFormat.Alignment = <%=iif(ind_alin_tit_col="C",wdAlignParagraphCenter,iif(ind_alin_tit_col="D",wdAlignParagraphRight,wdAlignParagraphLeft))%>
                          objWord.Selection.Cells.VerticalAlignment = <%=wdCellAlignVerticalCenter%>
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
													if "<%=ind_neg_tit_col%>" = "S" then
														objWord.Selection.Font.Bold = -1
													else
														objWord.Selection.Font.Bold = 0
													end if
                          objWord.Selection.TypeText "<%=strValor%>"
                          'Borda
                          <%
                                
                      ElseIf intCol <= objRs.Fields("qtd_nivel_linha") Then
                          'Cabeçalho da linha
                          strValor = ""
                          strValorTotalizador = ""
                          If Not objRs1.EOF Then
                              objRs1.Filter = "seq_linha = " & intLin - objRs.Fields("qtd_nivel_coluna")
                              If Not objRs1.EOF Then
                                  If (intCol = 1 And objRs.Fields("qtd_nivel_linha") = 1) Or _
                                          (intCol = 2 And objRs.Fields("qtd_nivel_linha") = 2) Or _
                                          (intCol = 3 And objRs.Fields("qtd_nivel_linha") = 3) Then
                                      strValor = objRs1.Fields("descr_nivel_interno") & ""
                                  ElseIf (intCol = 2 And objRs.Fields("qtd_nivel_linha") = 3) Then
                                      strValor = objRs1.Fields("descr_nivel_mediano") & ""
                                  ElseIf (intCol = 1 And objRs.Fields("qtd_nivel_linha") = 2) Or _
                                          (intCol = 1 And objRs.Fields("qtd_nivel_linha") = 3) Or _
                                          (intCol = 2 And objRs.Fields("qtd_nivel_linha") = 2) Or _
                                          (intCol = 3) Then
                                      strValor = objRs1.Fields("descr_nivel_externo") & ""
                                  End If
                                  If objRs1.Fields("ind_totalizador") & "" = "S" Then
                                      strValorTotalizador = " checked "
                                  End If
                              End If
                              objRs1.Filter = ""
                          End If
                          'strTabela = replace(strTabela, "<*" & intLin & "_" & intCol & "*>", strValor)
                          'strTabela = replace(strTabela, "|*" & intLin & "_" & qtdColunasTabela & "*|", strValorTotalizador)
                          %>
                          objWord.Selection.Shading.Texture = <%=wdTexture10Percent%>
                          objWord.Selection.ParagraphFormat.Alignment = <%=iif(ind_alin_tit_lin="C",wdAlignParagraphCenter,iif(ind_alin_tit_lin="D",wdAlignParagraphRight,wdAlignParagraphLeft))%>
                          objWord.Selection.Cells.VerticalAlignment = <%=wdCellAlignVerticalCenter%>
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
													if "<%=ind_neg_tit_lin%>" = "S" then
														objWord.Selection.Font.Bold = -1
													else
														objWord.Selection.Font.Bold = 0
													end if
                          objWord.Selection.TypeText "<%=strValor%>"
                          <%
                      ElseIf (intCol > objRs.Fields("qtd_nivel_linha") And intCol <= qtdColunasTabela) Or _
                              (intLin > objRs.Fields("qtd_nivel_coluna") And intLin <= qtdLinhasTabela) Then
                          'Detalhes
                          strValor = ""
                          If Not objRs3.EOF Then
                              objRs3.Filter = "seq_linha = " & intLin - objRs.Fields("qtd_nivel_coluna") & _
                                  " and seq_coluna = " & intCol - objRs.Fields("qtd_nivel_linha")
                              If Not objRs3.EOF Then
                                  strValor = objRs3.Fields("num_valor") & ""
                              End If
                              objRs3.Filter = ""
                          End If
                          'strTabela = replace(strTabela, "<*" & intLin & "_" & intCol & "*>", strValor)
                          %>
                          objWord.Selection.Shading.Texture = <%=wdTextureNone%>
                          objWord.Selection.ParagraphFormat.Alignment = <%=iif(ind_alin_dados="C",wdAlignParagraphCenter,iif(ind_alin_dados="D",wdAlignParagraphRight,wdAlignParagraphLeft))%>
                          objWord.Selection.Cells.VerticalAlignment = <%=wdCellAlignVerticalCenter%>
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
													if "<%=ind_neg_dados%>" = "S" then
														objWord.Selection.Font.Bold = -1
													else
														objWord.Selection.Font.Bold = 0
													end if
                          objWord.Selection.TypeText "<%=strValor%>"
                          <%
                      End If

                  Next
              Next
              objRs1.Close
              Set objRs1 = Nothing
              objRs2.Close
              Set objRs2 = Nothing
              objRs3.Close
              Set objRs3 = Nothing
							'
							if ind_tipo_planilha = "A" or _
									ind_tipo_planilha = "C" or _
									ind_tipo_planilha = "D" or _
									ind_tipo_planilha = "G" then
								%>
								objWord.Selection.Tables(1).Columns(1).Delete
								<%
							end if
              %>
							'objWord.Selection.Tables(1).select
							'objWord.Selection.Tables(1).Rows.Alignment = <%=wdAlignRowCenter%>
							objWord.Selection.Tables(1).Select
							objWord.Selection.Tables(1).AutoFormat 16, True, True, True, True, True, False, True, False, True
							
              'NOVO - APÓS MONTAGEM DA TABELA,
              'VARRER COLUNAS PARA QUEBRA-LA EM DIVERSAS TABELAS
              blnSairLoop = True
                    
              Do While blnSairLoop
                  For intCol = 1 To objWord.Selection.Tables(1).Columns.Count Step 8
                      'intCol varia 1, 9, 17, etc...
                      'Para cada bloco de tabelas ....
                      If objWord.Selection.Tables(1).Columns.Count > 8 Then
                          'caso exceda o tamanho máximo de cols na página,
                          'recortar o restante para montar nova tabela
                          'Selecionar
                          objWord.Selection.Tables(1).Cell(1, 9).Select
                          objWord.Selection.EndKey <%=wdRow%>, <%=wdExtend%>
                          objWord.Selection.EndOf <%=wdColumn%>, <%=wdExtend%>
                          'Recortar
                          objWord.Selection.Cut
                          'Reorganiza as colunas da tabela
                          <%ReorganizaLabelColunasTabela objRs%>
                          'Adicionar parágrafo
                          objWord.Selection.MoveDown <%=wdLine%>, 100
                          objWord.Selection.TypeParagraph

                          'Copiar
                          objWord.Selection.Paste

                      Else
                          blnSairLoop = False
                          'Reorganiza as colunas da tabela
                          <%ReorganizaLabelColunasTabela objRs%>
                      End If
                  Next
              Loop
              'objWord.Selection.MoveDown <%=wdLine%>, 100
              objWord.Selection.EndKey <%=wdStory%>
              objWord.Selection.TypeParagraph
              <%
          End If
          objRs.Close
          Set objRs = Nothing
          '
          objRsPrincipal.MoveNext
      Loop
    End If
  Else
    'Não é o status inicial da página
    'Se estiver no estado inicial da página,
    'pegar os dados do banco de dados, se houver
  End If
  objRsPrincipal.Close
  Set objRsPrincipal = Nothing
End Sub

Sub ReorganizaLabelColunasTabela(ByVal objRs)
%>
  'Dim lngColCabec
  'Dim lngLinCabec
  'Dim strTextoAnterior
  'Dim strTextoAnteriorOutroNivel
  'Dim lngColAnteiror
  'Dim lngColMarge
  'Dim blnSelecionarUltimaLinha
  On Error Resume Next
  <%
  If objRs.Fields("qtd_nivel_coluna") > 1 Then
  %>
      For lngLinCabec = 1 To <%=objRs.Fields("qtd_nivel_coluna") - 1%>
          'Para cada nível no cabeçalho da coluna
            
          strTextoAnterior = ""
          lngColAnteiror = 1
          lngColMarge = 0
          For lngColCabec = 1 To objWord.Selection.Tables(1).Columns.Count
              'Para cada nível no cabeçalho da coluna
              If lngLinCabec = 1 Or (lngLinCabec = 2 And <%=objRs.Fields("qtd_nivel_coluna")%> = 3) Then
                  'Primeiro nível mesclar independente dos demais níveis
                  objWord.Selection.Tables(1).Cell(lngLinCabec, lngColCabec - lngColMarge).Select
                  If Mid(objWord.Selection.Text, 1, Len(objWord.Selection.Text) - 2) <> "" Then
                      'Há algo na coluna
                      If strTextoAnterior = "" Then
                          'primeira coluna
                          strTextoAnterior = Mid(objWord.Selection.Text, 1, Len(objWord.Selection.Text) - 2)
                          lngColAnteiror = lngColCabec
                      Else
                          'Não é primeira coluna
                          If (UCase(strTextoAnterior) <> UCase(Mid(objWord.Selection.Text, 1, Len(objWord.Selection.Text) - 2))) Or _
                              lngColCabec = objWord.Selection.Tables(1).Columns.Count Then
                              'Mudou texto
                              'Selecionar sequencia de celulas para marge
                              If (UCase(strTextoAnterior) <> UCase(Mid(objWord.Selection.Text, 1, Len(objWord.Selection.Text) - 2))) Then
                                  blnSelecionarUltimaLinha = False
                              Else
                                  blnSelecionarUltimaLinha = True
                              End If
                              objWord.Selection.Tables(1).Cell(lngLinCabec, lngColAnteiror - lngColMarge).Select
                              For intLin = lngColAnteiror - lngColMarge To lngColCabec - lngColMarge + iif(lngColCabec = objWord.Selection.Tables(1).Columns.Count And blnSelecionarUltimaLinha = True, -1, -2) 'iif(lngColCabec <> objWord.Selection.Tables(1).Columns.Count, -2, -1)
                                  objWord.Selection.EndKey <%=wdLine%>, <%=wdExtend%>
                                  'lngColMarge = lngColMarge + 1 'decrementa as colunas
                                  'lngColCabec = lngColCabec - 1 'decrementa as colunas
                              Next
                              lngColMarge = lngColMarge + (lngColCabec - lngColMarge + iif(lngColCabec <> objWord.Selection.Tables(1).Columns.Count, -1, 0)) - (lngColAnteiror - lngColMarge)
                              'Mesclar Celulas
                              objWord.Selection.Cells.Merge
                              'Carrega texto de uma celula
                              objWord.Selection.TypeText strTextoAnterior
                              'Receleciona a coluna para prosseguir
                              If lngColCabec <> objWord.Selection.Tables(1).Columns.Count Then
                                  objWord.Selection.MoveRight <%=wdCharacter%>, 1
                              End If
                              objWord.Selection.EndKey <%=wdLine%>, <%=wdExtend%>
                              'Reposiciona valores anteriores
                              strTextoAnterior = Mid(objWord.Selection.Text, 1, Len(objWord.Selection.Text) - 2)
                              lngColAnteiror = lngColCabec
                          End If
                            
                      End If
                  End If
              End If
          Next
            
      Next
  <%
  End If
End Sub
%>
