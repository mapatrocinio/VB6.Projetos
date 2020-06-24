<%@ Language=VBScript %>

<!-- #include file="Includes/ValidaSessao.asp" -->
<!-- #include file="Includes/Funcoes.asp" -->

<!--#include file="Includes/ConexaobancoDados.asp"-->
<!--#include file="Includes/TrataErro.asp"-->
<!--#include file="Includes/ConstExcel.asp"-->
<%


   'Change HTML header to specify Excel's MIME content type.
   Response.Buffer = true
   'Response.ContentType = "application/msword"

  
  Dim objWordDoc
  Dim objWord
  Dim WordDocPath
  Dim strNomeArquivo

  'Formatar e carregar arquivo word
  'Relatório RAE
  Dim sql
  Dim sig_orgao_processo
  Dim seq_processo
  Dim ano_processo
  Dim objRs
  Dim objRsItens
  Dim strDescricao
  Dim strDescricao1
  '
  
	sig_orgao_processo	= request("Parametro_Sub_1")
	seq_processo		= request("Parametro_Sub_2")
	ano_processo		= request("Parametro_Sub_3")
  
  Set objWord = CreateObject("Word.Application")
  Set objWordDoc = CreateObject("Word.Document")
  
  Set objWordDoc = objWord.Documents.Add()

	sig_orgao_processo	= request("Parametro_Sub_1")
	seq_processo		= request("Parametro_Sub_2")
	ano_processo		= request("Parametro_Sub_3")

  strNomeArquivo = sig_orgao_processo & seq_processo & ano_processo & ".doc"
  
  'objWord.Visible = True
  'objWord.Selection.TypeText "arnaldo da silva sauro"
  'objWordDoc.PrintOut  True
  'objWordDoc.PrintPreview
	'MONTAR RELATÓRIO RAE
  Set objRs = Server.CreateObject("ADODB.Recordset")
  Set objRsItens = Server.CreateObject("ADODB.Recordset")

  sql = "spRelAuditoria '" & sig_orgao_processo & _
              "', '" & seq_processo & _
              "', '" & ano_processo & "'"
    
  objRs.Open sql, cn, 0, 1
  If objRs.EOF Then
    Response.Write "Ocorreu um erro na abertura do relatório !"
    Response.End
  End If
  sql = "spRelAuditoriaItens '" & sig_orgao_processo & _
              "', '" & seq_processo & _
              "', '" & ano_processo & "'"
    
  objRsItens.Open sql, cn, 0, 1
    
  If objRsItens.EOF Then
    Response.Write "Ocorreu um erro na abertura do relatório !"
    Response.End
  End If
  objWord.Visible=true
  'PRIMEIRA PÁGINA
  CargaPrimeiraPagina objRs, _
											objWordDoc, _
											objWord

  'CORPO DO DOCUMENTO
  CargaPaginasCorpo cn, _
                    objRsItens, _
                    sig_orgao_processo, _
                    seq_processo, _
                    ano_processo, _
										objWordDoc, _
										objWord

  'CABEÇALHO
  'CargaCabecalhoRodape objRsItens, _
	'									objWordDoc, _
	'									objWord
  'CARREGA NÚMERO DE PÁGINAS PRIMEIRA PÁGINA
  'CargaNumPagPriPag objWordDoc, _
	'									objWord
	

	
  ' setar a pasta para guardar os documentos gerados
  strWordDocPath = server.MapPath("Documentos") & "\" & strNomeArquivo
  objWordDoc.SaveAs strWordDocPath
  
  objWordDoc.Close  0
  objWord.Quit
  '

  Set objWordDoc = Nothing
  Set objWord = Nothing


  objRs.Close
  objRsItens.Close
  
  Set objRs = Nothing
  Set objRsItens = Nothing
	

	Response.redirect "Documentos/" & strNomeArquivo

'	Dim vntStream
'
'	Set objSDA = Server.CreateObject("bd_SDA.clsBD")
'	vntStream = objSDA.ReadBinFile("D:\Projetos\SDA\Site\ImpressaoHtml\RAE.doc")
'
'	Response.BinaryWrite(vntStream)
'
'	Set objSDA = Nothing
'
'	Response.End



' Set objWord = CreateObject("Word.Application")
' With objWord
'    ' Make the application visible.
'    .Visible = True
'    ' Open the document.
'	.Documents.Open ("D:\Projetos\SDA\Site\ImpressaoHtml\RAE.doc")
'	
' End With
' 
' Set objWord = Nothing7

Sub CargaPrimeiraPagina(ByVal objRs, _
												objWordDoc, _
												objWord)
  Dim strDescricao
  Dim strDescricao1
  
  ' --------- PRIMEIRA CAIXA - LOGO
  
  objWordDoc.Tables.Add objWord.Selection.Range, 1, 2
  objWord.Selection.InlineShapes.AddPicture server.MapPath("Imagens") & "\LogoInmetroRAE.jpg", False, True
  objWord.Selection.Tables(1).Columns(1).SetWidth 90, wdAdjustNone

  objWord.Selection.Tables(1).Columns(2).Select
  objWord.Selection.TypeText "RELATÓRIO DE AUDITORIA EXTRAORDINÁRIA ADMINISTRATIVA, CONTÁBIL E FINANCEIRA - RAE"
  objWord.Selection.Tables(1).Columns(2).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 12
  objWord.Selection.Font.Bold = wdToggle
  objWord.Selection.Tables(1).Columns(2).SetWidth 360, wdAdjustNone

  'SEGUNDA CAIXA
  objWord.Selection.MoveDown wdLine, 1
  objWord.Selection.TypeParagraph
  objWordDoc.Tables.Add objWord.Selection.Range, 1, 4
  
  objWord.Selection.Tables(1).Rows(1).SetHeight 27, wdRowHeightAtLeast
  'Primeira Coluna
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.TypeText "PROCESSO AUDIN" & Chr(10) & objRs.Fields("ProcessoAudin")
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(1).SetWidth 155, wdAdjustNone
  'Segunda Coluna
  objWord.Selection.Tables(1).Columns(2).Select
  objWord.Selection.TypeText "PERÍODO DA AUDITORIA" & Chr(10) & objRs.Fields("datainicio")
  objWord.Selection.Tables(1).Columns(2).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(2).SetWidth 155, wdAdjustNone
  'Terceira Coluna
  objWord.Selection.Tables(1).Columns(3).Select
  objWord.Selection.TypeText "DATA" & Chr(10) & Date & " - " & objRs.Fields("datafim")
  objWord.Selection.Tables(1).Columns(3).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(3).SetWidth 90, wdAdjustNone
  'Quarta Coluna
  objWord.Selection.Tables(1).Columns(4).Select
  objWord.Selection.Tables(1).Columns(4).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(4).SetWidth 50, wdAdjustNone
  'TERCEIRA CAIXA
  objWord.Selection.MoveDown wdLine, 1
  objWord.Selection.TypeParagraph
  objWordDoc.Tables.Add objWord.Selection.Range, 1, 1
  objWord.Selection.Tables(1).Rows(1).SetHeight 27, wdRowHeightAtLeast
    
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.TypeText "ÓRGÃO AUDITADO" & Chr(10) & objRs.Fields("orgaoauditado")
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(1).SetWidth 450, wdAdjustNone
  'QUARTA CAIXA
  objWord.Selection.MoveDown wdLine, 1
  objWord.Selection.TypeParagraph
  objWordDoc.Tables.Add objWord.Selection.Range, 1, 1
  objWord.Selection.Tables(1).Rows(1).SetHeight 17, wdRowHeightAtLeast
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.TypeText "EQUIPE AUDITORA"
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(1).SetWidth 450, wdAdjustNone
  objWord.Selection.MoveDown wdLine, 1
  objWord.Selection.TypeParagraph
  objWordDoc.Tables.Add objWord.Selection.Range, 1, 2
  objWord.Selection.Tables(1).Rows(1).SetHeight 17, wdRowHeightAtLeast
  'Primeira Coluna
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.TypeText "NOME"
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(1).SetWidth 225, wdAdjustNone
  'Segunda Coluna
  objWord.Selection.Tables(1).Columns(2).Select
  objWord.Selection.TypeText "UNIDADE"
  objWord.Selection.Tables(1).Columns(2).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(2).SetWidth 225, wdAdjustNone
  objWord.Selection.MoveUp wdLine, 1
  objWord.Selection.Delete wdCharacter, 1
  strDescricao = ""
  strDescricao1 = ""
  Do While Not objRs.EOF
    strDescricao = strDescricao & Chr(10) & UCase(objRs.Fields("equipeauditora"))
    strDescricao1 = strDescricao1 & Chr(10) & objRs.Fields("sig_uo_lotacao")
    '
    objRs.MoveNext
  Loop
  strDescricao = strDescricao & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10)
  strDescricao1 = strDescricao1 & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10)
  objRs.MoveFirst
  objWord.Selection.MoveDown wdLine, 1
  objWord.Selection.TypeParagraph
  objWordDoc.Tables.Add objWord.Selection.Range, 1, 2
  objWord.Selection.Tables(1).Rows(1).SetHeight 120, wdRowHeightAtLeast
  'Primeira Coluna
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.TypeText strDescricao
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(1).SetWidth 225, wdAdjustNone
  'Segunda Coluna
  objWord.Selection.Tables(1).Columns(2).Select
  objWord.Selection.TypeText strDescricao1
  objWord.Selection.Tables(1).Columns(2).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(2).SetWidth 225, wdAdjustNone
  objWord.Selection.MoveUp wdLine, 1
  objWord.Selection.Delete wdCharacter, 1
  'QUINTA CAIXA
  objWord.Selection.MoveDown wdLine, 100
  objWord.Selection.TypeParagraph
  objWordDoc.Tables.Add objWord.Selection.Range, 1, 1
  objWord.Selection.Tables(1).Rows(1).SetHeight 17, wdRowHeightAtLeast
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.TypeText "DETERMINAÇÃO DA AUDITORIA"
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(1).SetWidth 450, wdAdjustNone
  objWord.Selection.MoveDown wdLine, 1
  objWord.Selection.TypeParagraph
  objWordDoc.Tables.Add objWord.Selection.Range, 1, 1
  objWord.Selection.Tables(1).Rows(1).SetHeight 60, wdRowHeightAtLeast
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.TypeText "Oficio nº  " & objRs.Fields("numoficio") & " / Audin, de " & objRs.Fields("dataoficio")
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(1).SetWidth 450, wdAdjustNone
  objWord.Selection.MoveUp wdLine, 1
  objWord.Selection.Delete wdCharacter, 1
  'SEXTA CAIXA
  objWord.Selection.MoveDown wdLine, 100
  objWord.Selection.TypeParagraph
  objWordDoc.Tables.Add objWord.Selection.Range, 1, 1
  objWord.Selection.Tables(1).Rows(1).SetHeight 17, wdRowHeightAtLeast
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.TypeText "RECOMENDAÇÃO AO AUDITADO"
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(1).SetWidth 450, wdAdjustNone
  objWord.Selection.MoveDown wdLine, 1
  objWord.Selection.TypeParagraph
  objWordDoc.Tables.Add objWord.Selection.Range, 1, 1
  objWord.Selection.Tables(1).Rows(1).SetHeight 27, wdRowHeightAtLeast
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.TypeText iif(objRs.Fields("flag") = "1", "X", "   ") & "  SIM - PARA PROVIDÊNCIAS E/OU JUSTIFICATIVAS - 30 DIAS DO RECEBIMENTO DO RELATÓRIO" & Chr(10) & iif(objRs.Fields("flag") = "1", "   ", "X") & " NÃO"
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(1).SetWidth 450, wdAdjustNone
  objWord.Selection.MoveUp wdLine, 1
  objWord.Selection.Delete wdCharacter, 1
  'SÉTIMA CAIXA
  objWord.Selection.MoveDown wdLine, 100
  objWord.Selection.TypeParagraph
  objWordDoc.Tables.Add objWord.Selection.Range, 1, 1
  objWord.Selection.Tables(1).Rows(1).SetHeight 17, wdRowHeightAtLeast
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.TypeText "DE ACORDO / ENCAMINHAMENTO"
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(1).SetWidth 450, wdAdjustNone
  objWord.Selection.MoveDown wdLine, 1
  objWord.Selection.TypeParagraph
  objWordDoc.Tables.Add objWord.Selection.Range, 1, 1
  objWord.Selection.Tables(1).Rows(1).SetHeight 50, wdRowHeightAtLeast
        
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalTop
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(1).SetWidth 450, wdAdjustNone
  objWord.Selection.MoveUp wdLine, 1
  objWord.Selection.Delete wdCharacter, 1
  objWord.Selection.TypeText objRs.Fields("encaminhamento") & ""
  objWord.Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    
  objWord.Selection.MoveDown wdLine, 1
  objWord.Selection.TypeParagraph
  objWordDoc.Tables.Add objWord.Selection.Range, 1, 1
  objWord.Selection.Tables(1).Rows(1).SetHeight 40, wdRowHeightAtLeast
        
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(1).SetWidth 450, wdAdjustNone
  objWord.Selection.MoveUp wdLine, 1
  objWord.Selection.Delete wdCharacter, 1
  objWord.Selection.TypeText "__________________________________________________________" & Chr(10) & objRs.Fields("nomechefe") & Chr(10) & objRs.Fields("crcchefe") & Chr(10) & objRs.Fields("funcchefe")
  objWord.Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    
  objWord.Selection.MoveDown wdLine, 100
  'objWord.Selection.TypeParagraph
  objWord.Selection.InsertBreak
    
End Sub


Sub CargaPaginasCorpo(ByVal objConn, _
                      ByVal objRsItens, _
                      ByVal sig_orgao_processo, _
                      ByVal seq_processo, _
                      ByVal ano_processo, _
											objWordDoc, _
											objWord)
    
  Dim seq_relatorio
  Dim status_detalhes
  Dim ind_tabela
    

  'PRIMEIRO ITEM
  objWord.Selection.MoveDown wdLine, 1
  objWord.Selection.TypeParagraph
  objWord.Selection.TypeText objRsItens.Fields("Apresentacao") & ""
  objWord.Selection.MoveUp wdParagraph, 1, wdExtend
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
    
  'SEGUNDO ITEM
  objWord.Selection.MoveDown wdLine, 1
  objWord.Selection.TypeParagraph
  objWord.Selection.TypeParagraph
  objWord.Selection.TypeText "1. Introdução"
  objWord.Selection.MoveUp wdParagraph, 1, wdExtend
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 12
  objWord.Selection.Font.Bold = wdToggle

  objWord.Selection.MoveDown wdLine, 1
  objWord.Selection.TypeParagraph
  objWord.Selection.TypeParagraph
  objWord.Selection.TypeText objRsItens.Fields("Introducao") & ""
  objWord.Selection.MoveUp wdParagraph, 1, wdExtend
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Font.Bold = wdToggle
    
  'TERCEIRO ITEM
  objWord.Selection.MoveDown wdLine, 1
  objWord.Selection.TypeParagraph
  objWord.Selection.TypeParagraph
  objWord.Selection.TypeText "2. Dos Exames realizados"
  objWord.Selection.MoveUp wdParagraph, 1, wdExtend
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 12
  objWord.Selection.Font.Bold = wdToggle

  objWord.Selection.MoveDown wdLine, 1
  objWord.Selection.TypeParagraph
  objWord.Selection.TypeParagraph
  objWord.Selection.TypeText objRsItens.Fields("Exames") & ""
  objWord.Selection.MoveUp wdParagraph, 1, wdExtend
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Font.Bold = wdToggle

  objWord.Selection.MoveDown wdLine, 1
  objWord.Selection.TypeParagraph
  objWord.Selection.TypeParagraph
    
  Do While Not objRsItens.EOF
    'TERCEIRO ITEM
        
    objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    objWord.Selection.Font.Name = "Arial"
    objWord.Selection.Font.Size = 12
    objWord.Selection.TypeText objRsItens.Fields("descr_relatorio") & ""
    '
    'carregar seq_relatório
    seq_relatorio = iif(objRsItens.Fields("seq_relatorio") & "" = "", 0, objRsItens.Fields("seq_relatorio"))
    'Carregar planilha
    status_detalhes = "inicial"
    ind_tabela = "relatorio_salvar_listar"
    exibir_planilha_item objConn, _
                         seq_relatorio, _
                         status_detalhes, _
                         ind_tabela, _
                         sig_orgao_processo, _
                         seq_processo, _
                         ano_processo, _
												 objWordDoc, _
												 objWord

    'Incrementa Contador
    objRsItens.MoveNext
  Loop
  objRsItens.MoveFirst

End Sub


Sub exibir_planilha_item(ByVal objConn, _
                         ByVal seq_relatorio, _
                         ByVal status_detalhes, _
                         ByVal ind_tabela, _
                         ByVal sig_orgao_processo, _
                         ByVal seq_processo, _
                         ByVal ano_processo, _
												 objWordDoc, _
												 objWord)
    
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

  Dim blnSairLoop

  Set objRsPrincipal = CreateObject("ADODB.Recordset")

  'If seq_relatorio <> 4 Then Exit Sub
  Select Case ind_tabela
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

  End Select
  objRsPrincipal.Open strSql, objConn, 0, 1
    
  If status_detalhes = "inicial" Then
    'Se estiver no estado inicial da página,
    'pegar os dados do banco de dados, se houver

    strTabela = ""
    If Not objRsPrincipal.EOF Then
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
              'Definidos as linhas/colunas da tabela monta a tabela
              'em variável
              objWord.Selection.MoveDown wdLine, 1
              objWord.Selection.TypeParagraph
                    
              objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
              objWord.Selection.Font.Name = "Arial"
              objWord.Selection.Font.Size = 12
              objWord.Selection.TypeText descr_titulo
                    
              objWord.Selection.MoveDown wdLine, 1
              objWord.Selection.TypeParagraph
                    
              objWordDoc.Tables.Add objWord.Selection.Range, 1, 1
              objWord.Selection.Tables(1).Columns(1).SetWidth 60, wdAdjustProportional
              'Seleciona Coluna
              objWord.Selection.Tables(1).Columns(1).Select
                    
              'On Error Resume Next
                    
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
                      objWord.Selection.Tables(1).Rows.Add
                  End If
                  For intCol = 1 To qtdColunasTabela
                      If intCol <> 1 And intLin = 1 Then
                          objWord.Selection.Tables(1).Columns.Add
                      End If
                      objWord.Selection.Tables(1).Cell(intLin, intCol).Select
                        
                      If (intCol <= objRs.Fields("qtd_nivel_linha") And intLin <= objRs.Fields("qtd_nivel_coluna")) Then 'Or _
                              '(intCol = qtdColunasTabela And intLin = qtdLinhasTabela) Then
                          'interseção entre linha x coluna do cabeçalho
                          'da tabela não adiciona nada
                          'strTabela = "" 'replace(strTabela, "<*" & intLin & "_" & intCol & "*>", "")
                          'Não faz nada
'''                                objWord.Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
'''                                objWord.Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
'''                                objWord.Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
'''                                objWord.Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
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
                          objWord.Selection.Shading.Texture = wdTexture10Percent
                          objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                          objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
                          objWord.Selection.TypeText strValor
                          'Borda
'''                                With objWord.Selection.Borders(wdBorderTop)
'''                                    .LineStyle = Options.DefaultBorderLineStyle
'''                                    .LineWidth = Options.DefaultBorderLineWidth
'''                                    .ColorIndex = Options.DefaultBorderColorIndex
'''                                End With
'''                                With objWord.Selection.Borders(wdBorderLeft)
'''                                    .LineStyle = Options.DefaultBorderLineStyle
'''                                    .LineWidth = Options.DefaultBorderLineWidth
'''                                    .ColorIndex = Options.DefaultBorderColorIndex
'''                                End With
'''                                With objWord.Selection.Borders(wdBorderBottom)
'''                                    .LineStyle = Options.DefaultBorderLineStyle
'''                                    .LineWidth = Options.DefaultBorderLineWidth
'''                                    .ColorIndex = Options.DefaultBorderColorIndex
'''                                End With
'''                                With objWord.Selection.Borders(wdBorderRight)
'''                                    .LineStyle = Options.DefaultBorderLineStyle
'''                                    .LineWidth = Options.DefaultBorderLineWidth
'''                                    .ColorIndex = Options.DefaultBorderColorIndex
'''                                End With
                                
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
                          objWord.Selection.Shading.Texture = wdTexture10Percent
                          objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                          objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
                          objWord.Selection.TypeText strValor
'''                                'Borda
'''                                With objWord.Selection.Borders(wdBorderTop)
'''                                    .LineStyle = Options.DefaultBorderLineStyle
'''                                    .LineWidth = Options.DefaultBorderLineWidth
'''                                    .ColorIndex = Options.DefaultBorderColorIndex
'''                                End With
'''                                With objWord.Selection.Borders(wdBorderLeft)
'''                                    .LineStyle = Options.DefaultBorderLineStyle
'''                                    .LineWidth = Options.DefaultBorderLineWidth
'''                                    .ColorIndex = Options.DefaultBorderColorIndex
'''                                End With
'''                                With objWord.Selection.Borders(wdBorderBottom)
'''                                    .LineStyle = Options.DefaultBorderLineStyle
'''                                    .LineWidth = Options.DefaultBorderLineWidth
'''                                    .ColorIndex = Options.DefaultBorderColorIndex
'''                                End With
'''                                With objWord.Selection.Borders(wdBorderRight)
'''                                    .LineStyle = Options.DefaultBorderLineStyle
'''                                    .LineWidth = Options.DefaultBorderLineWidth
'''                                    .ColorIndex = Options.DefaultBorderColorIndex
'''                                End With
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
                          objWord.Selection.Shading.Texture = wdTextureNone
                          objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                          objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
                          objWord.Selection.Font.Name = "Arial"
                          objWord.Selection.Font.Size = 9
                          objWord.Selection.TypeText strValor
'''                                'Borda
'''                                With objWord.Selection.Borders(wdBorderTop)
'''                                    .LineStyle = Options.DefaultBorderLineStyle
'''                                    .LineWidth = Options.DefaultBorderLineWidth
'''                                    .ColorIndex = Options.DefaultBorderColorIndex
'''                                End With
'''                                With objWord.Selection.Borders(wdBorderLeft)
'''                                    .LineStyle = Options.DefaultBorderLineStyle
'''                                    .LineWidth = Options.DefaultBorderLineWidth
'''                                    .ColorIndex = Options.DefaultBorderColorIndex
'''                                End With
'''                                With objWord.Selection.Borders(wdBorderBottom)
'''                                    .LineStyle = Options.DefaultBorderLineStyle
'''                                    .LineWidth = Options.DefaultBorderLineWidth
'''                                    .ColorIndex = Options.DefaultBorderColorIndex
'''                                End With
'''                                With objWord.Selection.Borders(wdBorderRight)
'''                                    .LineStyle = Options.DefaultBorderLineStyle
'''                                    .LineWidth = Options.DefaultBorderLineWidth
'''                                    .ColorIndex = Options.DefaultBorderColorIndex
'''                                End With
                      End If
                  Next
              Next
              objRs1.Close
              Set objRs1 = Nothing
              objRs2.Close
              Set objRs2 = Nothing
              objRs3.Close
              Set objRs3 = Nothing
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
                          objWord.Selection.EndKey wdRow, wdExtend
                          objWord.Selection.EndOf wdColumn, wdExtend
                          'Recortar
                          objWord.Selection.Cut
                          'Reorganiza as colunas da tabela
                          ReorganizaLabelColunasTabela objRs
                          'Adicionar parágrafo
                          objWord.Selection.MoveDown wdLine, 100
                          objWord.Selection.TypeParagraph
                          'Copiar
                          objWord.Selection.Paste

                      Else
                          blnSairLoop = False
                          'Reorganiza as colunas da tabela
                          ReorganizaLabelColunasTabela objRs
                      End If
                  Next
              Loop
              objWord.Selection.MoveDown wdLine, 100
              objWord.Selection.TypeParagraph
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
  Dim lngColCabec
  Dim lngLinCabec
  Dim strTextoAnterior
  Dim strTextoAnteriorOutroNivel
  Dim lngColAnteiror
  Dim lngColMarge
  Dim blnSelecionarUltimaLinha
  On Error Resume Next
  If objRs.Fields("qtd_nivel_coluna") > 1 Then
      For lngLinCabec = 1 To objRs.Fields("qtd_nivel_coluna") - 1
          'Para cada nível no cabeçalho da coluna
            
          strTextoAnterior = ""
          lngColAnteiror = 1
          lngColMarge = 0
          For lngColCabec = 1 To objWord.Selection.Tables(1).Columns.Count
              'Para cada nível no cabeçalho da coluna
              If lngLinCabec = 1 Or (lngLinCabec = 2 And objRs.Fields("qtd_nivel_coluna") = 3) Then
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
                                  objWord.Selection.EndKey wdLine, wdExtend
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
                                  objWord.Selection.MoveRight wdCharacter, 1
                              End If
                              objWord.Selection.EndKey wdLine, wdExtend
                              'Reposiciona valores anteriores
                              strTextoAnterior = Mid(objWord.Selection.Text, 1, Len(objWord.Selection.Text) - 2)
                              lngColAnteiror = lngColCabec
                          End If
                            
                      End If
                  End If
              End If
          Next
            
      Next
  End If

End Sub

    
Sub CargaCabecalhoRodape(ByVal objRsItens, _
												 objWordDoc, _
												 objWord)
  If objWord.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
    objWord.ActiveWindow.Panes(2).Close
  End If
  If objWord.ActiveWindow.ActivePane.View.Type = wdNormalView Or objWord.ActiveWindow. _
    ActivePane.View.Type = wdOutlineView Or objWord.ActiveWindow.ActivePane.View.Type _
     = wdMasterView Then
    objWord.ActiveWindow.ActivePane.View.Type = wdPageView
  End If
  'CABEÇALHO
  objWord.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
  'Apagar tudo
  objWord.Selection.WholeStory
  objWord.Selection.Delete wdCharacter, 1


  objWordDoc.Tables.Add objWord.Selection.Range, 1, 3
  objWord.Selection.Tables(1).Rows(1).SetHeight 27, wdRowHeightAtLeast
  'Primeira Coluna
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.TypeText "RELATÓRIO DE AUDITORIA EXTRAORDINÁRIA"
  objWord.Selection.Tables(1).Columns(1).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 12
  objWord.Selection.Font.Bold = wdToggle
  objWord.Selection.Tables(1).Columns(1).SetWidth 300, wdAdjustNone
  'Segunda Coluna
  objWord.Selection.Tables(1).Columns(2).Select
  objWord.Selection.TypeText "PROCESSO AUDIN" & Chr(10) & objRsItens.Fields("ProcessoAudin")
  objWord.Selection.Tables(1).Columns(2).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(2).SetWidth 100, wdAdjustNone
  'Terceira Coluna
  objWord.Selection.Tables(1).Columns(3).Select
  objWord.Selection.TypeText "PÁGINA" & Chr(10) '& "2/2"
  objWord.Selection.Fields.Add objWord.Selection.Range, wdFieldPage
  objWord.Selection.TypeText " / "
  objWord.Selection.Fields.Add objWord.Selection.Range, wdFieldNumPages


  objWord.Selection.Tables(1).Columns(3).Select
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 9
  objWord.Selection.Tables(1).Columns(3).SetWidth 50,  _
      wdAdjustNone
  'FIM CABEÇALHO
  'RODAPÉ
  objWord.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
  'Apagar tudo
  objWord.Selection.WholeStory
  objWord.Selection.Delete wdCharacter, 1
  'Adicionar rodapé
  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 8
  objWord.Selection.TypeText "FOR - AUDIN - 008 - Ver. 00 - Apr. JAN/03 - Pg. "
  objWord.Selection.Fields.Add objWord.Selection.Range, wdFieldPage
  objWord.Selection.TypeText " / "
  objWord.Selection.Fields.Add objWord.Selection.Range, wdFieldNumPages
  'Cabeçalho e rodapé da primeira página
  With objWordDoc.Sections(1)
    .PageSetup.DifferentFirstPageHeaderFooter = True
    .Headers(wdHeaderFooterFirstPage).Range.Text = ""
    .Footers(wdHeaderFooterFirstPage).Range.Text = ""
  End With
  objWord.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub

Sub CargaNumPagPriPag(objWordDoc, _
											objWord)
  objWord.Selection.HomeKey wdStory
  objWord.Selection.MoveDown wdLine, 2
  objWord.Selection.MoveRight wdCell
  objWord.Selection.MoveRight wdCell
  objWord.Selection.MoveRight wdCell
  objWord.Selection.TypeText "PÁGINA" & Chr(10)

  objWord.Selection.Fields.Add objWord.Selection.Range, wdFieldPage
  objWord.Selection.TypeText " / "
  objWord.Selection.Fields.Add objWord.Selection.Range, wdFieldNumPages

  objWordDoc.Browser.Next
  objWord.Selection.MoveUp wdLine, 1

  objWord.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
  objWord.Selection.Font.Name = "Arial"
  objWord.Selection.Font.Size = 8
  objWord.Selection.TypeText "FOR - AUDIN - 008 - Ver. 00 - Apr. JAN/03 - Pg. "
  objWord.Selection.Fields.Add objWord.Selection.Range, wdFieldPage
  objWord.Selection.TypeText " / "
  objWord.Selection.Fields.Add objWord.Selection.Range, wdFieldNumPages

End Sub


%>
