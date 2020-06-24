<%'cabeçalho da página dados_tabela.asp%>
<%
	select case ind_tabela
	case "template_relatorio"
		'Tabela de cabeçalho
		Acao =				Request("Acao")
		Aux				    = Request("Aux")
		NumProcesso         = Ucase(Request("NumProcesso"))
		sig_orgao_processo  = Request("sig_orgao_processo")
		seq_Processo		= Request("seq_Processo")
		ano_processo		= Request("ano_processo")
		seq_tipo_documento 	= Request("seq_tipo_documento")
		seq_topico_doc_conclu	= request("seq_topico_doc_conclu")
		seq_template 		= request("seq_template")
		%>
    <tr> 
			<td class="texto" width="20%">Processo:
			</td>
			<td class="texto" width="80%"><%=NumProcesso%></td>
    </tr>
    <tr><td class="texto">&nbsp;
		<input type="hidden" name="retorno" value="<%=retorno%>">
		<input type="hidden" name="Acao" value="<%=Acao%>">
		<input type="hidden" name="Aux">
		<input type="hidden" name="sig_orgao_processo" value="<%=sig_orgao_processo%>">
		<input type="hidden" name="seq_processo" value="<%=seq_processo%>">
		<input type="hidden" name="ano_processo" value="<%=ano_processo%>">
		<input type="hidden" name="seq_relatorio">
		<input type="hidden" name="ind_alteracao">
		<input type="hidden" name="seq_tipo_documento" value="<%=seq_tipo_documento%>">
		<input type="hidden" name="seq_topico_doc_conclu" value="<%=seq_topico_doc_conclu%>">
		<input type="hidden" name="seq_template" value="<%=seq_template%>">
    </td></tr>
		<%
	case "template"
		'Tabela de cabeçalho
		str_acao_template_cadastrar			= request("acao_template_cadastrar")
		str_controle_template_cadastrar = request("controle_template_cadastrar")
		str_seq_tipo_documento					= request("seq_tipo_documento")
		str_seq_topico_documento				= request("seq_topico_documento")
		str_hid_seq_template						= request("hid_seq_template")

		str_descr_template 				= request("descr_template")
		str_descr_tipo_documento 	= request("descr_tipo_documento")
		str_descr_conteudo 				= request("descr_conteudo")
		
		strSql = "select tipo_doc_auditoria.descr_tipo_documento, topico_doc_auditoria.descr_conteudo, template_doc_auditoria.descr_template " & _
			"from tipo_doc_auditoria " & _
			" inner join topico_doc_auditoria on tipo_doc_auditoria.seq_tipo_documento = topico_doc_auditoria.seq_tipo_documento " & _
			" inner join template_doc_auditoria on topico_doc_auditoria.seq_tipo_documento = template_doc_auditoria.seq_tipo_documento " & _
			" and topico_doc_auditoria.seq_topico_doc = template_doc_auditoria.seq_topico_doc " & _
			"where template_doc_auditoria.seq_tipo_documento = " & str_seq_tipo_documento & _
			" and template_doc_auditoria.seq_topico_doc = " & str_seq_topico_documento & _
			" and template_doc_auditoria.seq_template = " & str_hid_seq_template
		set objRs = server.CreateObject("ADODB.RECORDSET")
		objRs.Open strSql,cn,0,1
		if not objRs.EOF then
			descr_linha			= objRs.Fields("descr_tipo_documento").Value
			descr_linha_1			= objRs.Fields("descr_conteudo").Value
			descr_linha_2			= objRs.Fields("descr_template").Value
		end if
		objRs.Close
		set objRs = nothing
		'Escrever campos desabilitados e campos hidden 
		%>
    <tr> 
			<td class="texto" width="20%">Tipo de Documento:
			</td>
			<td class="texto" width="80%"><%=descr_linha%></td>
    </tr>
    <tr> 
			<td class="texto" width="20%">Tópico Documento:
			</td>
			<td class="texto" width="80%"><%=descr_linha_1%></td>
    </tr>
    <tr> 
			<td class="texto" width="20%">Template:
			</td>
			<td class="texto" width="80%"><%=descr_linha_2%></td>
    </tr>
    <tr><td class="texto">&nbsp;
		<input type="hidden" name="acao_template_cadastrar" value="<%=str_acao_template_cadastrar%>">
		<input type="hidden" name="controle_template_cadastrar"  value="<%=str_controle_template_cadastrar%>">
		
		<input type="hidden" name="descr_template"  value="<%=str_descr_template%>">
		<input type="hidden" name="descr_tipo_documento"  value="<%=str_descr_tipo_documento%>">
		<input type="hidden" name="descr_conteudo"  value="<%=str_descr_conteudo%>">
	
		<!-- #include file="template_listar_hidden.asp"-->
			
    </td></tr>
		<%
	case "relatorio_auditoria"
		'Tabela de relatorio_auditoria
		Acao								= Request("Acao")
		NumProcesso					= Request("NumProcesso")
		sig_orgao_processo	= Request("Sig_Orgao_Processo")
		seq_Processo				= Request("Seq_Processo")
		ano_processo				= Request("Ano_Processo")
		seq_relatorio				= Request("seq_relatorio")
		seq_checklist				= Request("Seq_Checklist")
		seq_Comentario			= Request("Seq_Comentario")
		descr_linha					= replace(Request("txtLinha"), vbCrLf,"")
		
		strSql = "select relatorio_auditoria.descr_linha " & _
			"from relatorio_auditoria " & _
			"where relatorio_auditoria.ano_processo = " & ano_processo & _
			"and relatorio_auditoria.sig_orgao_processo = '" & sig_orgao_processo & "'" & _
			"and relatorio_auditoria.seq_processo = " & seq_processo & _
			"and relatorio_auditoria.seq_relatorio = " & seq_relatorio
			
		set objRs = server.CreateObject("ADODB.RECORDSET")
		objRs.Open strSql,cn,0,1
		if not objRs.EOF then
			descr_linha			= objRs.Fields("descr_linha").Value
		end if
		objRs.Close
		set objRs = nothing
		'Escrever campos desabilitados e campos hidden 
		%>
    <tr> 
			<td class="texto" width="20%">Ítem:
			</td>
			<td class="texto" width="80%"><%=descr_linha%></td>
    </tr>
    <tr><td class="texto">&nbsp;
			<input type="hidden" name="Acao" value="<%=Acao%>">
			<input type="hidden" name="sig_orgao_processo" value="<%=sig_orgao_processo%>">
			<input type="hidden" name="seq_processo" value="<%=seq_processo%>">
			<input type="hidden" name="ano_processo" value="<%=ano_processo%>">
			<input type="hidden" name="seq_relatorio" value="<%=seq_relatorio%>">
			<input type="hidden" name="NumProcesso" value="<%=NumProcesso%>">
			<input type="hidden" name="seq_checklist" value="<%=seq_checklist%>">
			<input type="hidden" name="seq_Comentario" value="<%=seq_Comentario%>">
			<input type="hidden" name="txtLinha" value="<%=descr_linha%>">
			
    </td></tr>
		<%
	case "relatorio_salvar_listar"
		'Tabela de relatorio_auditoria
		retorno										= Request("retorno")
		Acao											= Request("Acao")
		sig_orgao_processo				= Request("sig_orgao_processo")
		seq_processo							= Request("seq_processo")
		ano_processo							= Request("ano_processo")
		seq_relatorio							= Request("seq_relatorio")
		ind_alteracao							= Request("ind_alteracao")
		NumProcesso								= Request("NumProcesso")
		
		strSql = "select relatorio_auditoria.descr_linha " & _
			"from relatorio_auditoria " & _
			"where relatorio_auditoria.ano_processo = " & ano_processo & _
			"and relatorio_auditoria.sig_orgao_processo = '" & sig_orgao_processo & "'" & _
			"and relatorio_auditoria.seq_processo = " & seq_processo & _
			"and relatorio_auditoria.seq_relatorio = " & seq_relatorio
			
		set objRs = server.CreateObject("ADODB.RECORDSET")
		objRs.Open strSql,cn,0,1
		if not objRs.EOF then
			descr_linha			= objRs.Fields("descr_linha").Value
		end if
		objRs.Close
		set objRs = nothing
		'Escrever campos desabilitados e campos hidden 
		%>
    <tr> 
			<td class="texto" width="20%">Ítem:
			</td>
			<td class="texto" width="80%"><%=descr_linha%></td>
    </tr>
    <tr><td class="texto">&nbsp;
			<input type="hidden" name="retorno" value="<%=retorno%>">
			<input type="hidden" name="Acao" value="<%=Acao%>">
			<input type="hidden" name="sig_orgao_processo" value="<%=sig_orgao_processo%>">
			<input type="hidden" name="seq_processo" value="<%=seq_processo%>">
			<input type="hidden" name="ano_processo" value="<%=ano_processo%>">
			<input type="hidden" name="seq_relatorio" value="<%=seq_relatorio%>">
			<input type="hidden" name="ind_alteracao" value="<%=ind_alteracao%>">
			<input type="hidden" name="NumProcesso" value="<%=NumProcesso%>">
    </td></tr>
		<%
	case "pergunta"
		'Tabela de pergunta
		seq_Assunto		= request("SelAssunto")
		seq_Item_sa		= request("SelItemSa")
		strSql = "select item_sa.descr_item_sa, assunto.descr_assunto " & _
			"from item_sa inner join assunto on assunto.seq_assunto = item_sa.seq_assunto " & _
			"where item_sa.seq_assunto = " & seq_Assunto & _
			"and item_sa.seq_item_sa = " & seq_item_sa
		set objRs = server.CreateObject("ADODB.RECORDSET")
		objRs.Open strSql,cn,0,1
		if not objRs.EOF then
			descr_Assunto		= objRs.Fields("descr_assunto").Value
			descr_Item_sa		= objRs.Fields("descr_item_sa").Value
		end if
		objRs.Close
		set objRs = nothing
		'Escrever campos desabilitados e campos hidden 
		%>
    <tr> 
			<td class="texto" width="20%">Assunto:
			</td>
			<td class="texto" width="80%"><%=descr_Assunto%></td>
    </tr>
    <tr> 
			<td class="texto" width="20%">Pergunta:
			</td>
			<td class="texto" width="80%"><%=descr_Item_sa%></td>
    </tr>
    <tr><td class="texto">&nbsp;
		<input type="hidden" name="SelAssunto" value="<%=seq_Assunto%>">
		<input type="hidden" name="SelItemSa" value="<%=seq_Item_sa%>">
    </td></tr>
		<%
	case "paaai"
		'Tabela de paaai
		ano_paaai			= request("ano_busca")
		cod_uo				= request("unidade_busca")
		strSql = "select vw_uo_auditoria.num_identificacao, vw_uo_auditoria.sig_uo " & _
			"from paaai left join vw_uo_auditoria on paaai.cod_uo = vw_uo_auditoria.cod_uo " & _
			"where paaai.cod_uo = '" & cod_uo & _
			"' and paaai.ano_paaai = " & ano_paaai
		set objRs = server.CreateObject("ADODB.RECORDSET")
		objRs.Open strSql,cn,0,1
		if not objRs.EOF then
			sig_uo							= objRs.Fields("sig_uo").Value
			num_identificacao		= objRs.Fields("num_identificacao").Value
			
		end if
		objRs.Close
		set objRs = nothing
		'Escrever campos desabilitados e campos hidden 
		%>
    <tr> 
			<td class="texto" width="20%">Ano:
			</td>
			<td class="texto" width="80%"><%=ano_paaai%></td>
    </tr>
    <tr> 
			<td class="texto" width="20%">Unidade:
			</td>
			<td class="texto" width="80%"><%=trim(num_identificacao & "") & " - " & trim(sig_uo & "")%></td>
    </tr>
    <tr><td class="texto">&nbsp;
		<input type="hidden" name="ano_busca" value="<%=ano_paaai%>">
		<input type="hidden" name="unidade_busca" value="<%=cod_uo%>">
		<input type="hidden" name="aux" value="Buscar">
    </td></tr>
		<%
	case "sa_item_auditoria"
		'Tabela de sa_item_auditoria
		ano_processo					= Request("Ano_Processo")
		sig_orgao_processo		= Request("Sig_Orgao_Processo")
		seq_Processo					= Request("Seq_Processo")
		seq_sa								= Request("Seq_Sa")
		seq_sa_complementar		= Request("seq_sa_complementar")
		seq_assunto						= Request("Seq_Assunto")
		seq_item_sa						= Request("SelItemSa")

		NumProcesso						= Ucase(Request("NumProcesso"))
		seq_area							= Request("Seq_Area")
		acao									= Request("Acao")
		ListaItens						= Request("ListaItens")
		
		strSql = "select CASE WHEN not descr_item_modificado_sa  is null THEN descr_item_modificado_sa ELSE descr_item_sa END as descr_item_sa " & _
			"from sa_item_auditoria inner join item_sa " & _
			"on sa_item_auditoria.seq_item_sa = item_sa.seq_item_sa " & _
			" and sa_item_auditoria.seq_assunto = item_sa.seq_assunto " & _
			"where ano_processo = " & ano_processo & _
			" and sig_orgao_processo = '" & sig_orgao_processo & "'" & _
			" and seq_Processo = " & seq_Processo & _
			" and seq_sa = " & seq_sa & _
			" and seq_sa_complementar = '" & seq_sa_complementar & "'" & _
			" and sa_item_auditoria.seq_assunto = " & seq_assunto & _
			" and sa_item_auditoria.seq_item_sa = " & seq_item_sa

		set objRs = server.CreateObject("ADODB.RECORDSET")
		objRs.Open strSql,cn,0,1
		if not objRs.EOF then
			descr_item_sa						= objRs.Fields("descr_item_sa").Value
		end if
		objRs.Close
		set objRs = nothing
		'Escrever campos desabilitados e campos hidden 
		%>
    <tr> 
			<td class="texto" width="20%">Número do Processo:
			</td>
			<td class="texto" width="80%"><%=NumProcesso%></td>
    </tr>
    <tr> 
			<td class="texto" width="20%">Ítem da SA:
			</td>
			<td class="texto" width="80%"><%=descr_item_sa%></td>
    </tr>
    <tr><td class="texto">&nbsp;
    
		<input type="hidden" name="Ano_Processo" value="<%=ano_processo%>">
		<input type="hidden" name="Sig_Orgao_Processo" value="<%=sig_orgao_processo%>">
		<input type="hidden" name="Seq_Processo" value="<%=seq_Processo%>">
		<input type="hidden" name="Seq_Sa" value="<%=seq_sa%>">
		<input type="hidden" name="seq_sa_complementar" value="<%=seq_sa_complementar%>">
		<input type="hidden" name="Seq_Assunto" value="<%=seq_assunto%>">
		<input type="hidden" name="SelItemSa" value="<%=seq_item_sa%>">
		<input type="hidden" name="NumProcesso" value="<%=NumProcesso%>">
		<input type="hidden" name="Seq_Area" value="<%=seq_area%>">
		<input type="hidden" name="Acao" value="<%=acao%>">
		<input type="hidden" name="ListaItens" value="<%=ListaItens%>">
    </td></tr>
		<%
	case "comentario_item_sa"
		'Tabela de comentario_item_sa
		Acao									= Request("Acao")
		NumProcesso						= Request("NumProcesso")
		sig_orgao_processo		= Request("sig_orgao_processo")
		seq_Processo					= Request("seq_Processo")
		ano_processo					= Request("ano_processo")
		seq_sa								= Request("seq_sa")
		seq_sa_complementar		= Request("seq_sa_complementar")
		seq_area							= Request("seq_area")
		seq_assunto						= Request("seq_assunto")
		seq_item_sa						= Request("seq_item_sa")
		seq_checklist					= Request("seq_checklist")
		seq_Comentario				= Request("seq_Comentario")
		descr_Comentario			= Request("txtComentario")
		descr_anexo						= Request("txtAnexo")
		descr_Recomendacao		= Request("txtRecomendacao")
		ind_relatorio					= Request("chk_relatorio")
		ind_ata								= Request("chk_ata")
		Origem								= Request("Origem")
		retorno								= Request("retorno")
		ListaItensSa					= Request("ListaItensSa")

		strSql = "select CASE WHEN not descr_item_modificado_sa  is null THEN descr_item_modificado_sa ELSE descr_item_sa END as descr_item_sa " & _
			"from sa_item_auditoria inner join item_sa " & _
			"on sa_item_auditoria.seq_item_sa = item_sa.seq_item_sa " & _
			" and sa_item_auditoria.seq_assunto = item_sa.seq_assunto " & _
			"where ano_processo = " & ano_processo & _
			" and sig_orgao_processo = " & sig_orgao_processo & _
			" and seq_Processo = " & seq_Processo & _
			" and seq_sa = " & seq_sa & _
			" and seq_sa_complementar = " & seq_sa_complementar & _
			" and sa_item_auditoria.seq_assunto = " & seq_assunto & _
			" and sa_item_auditoria.seq_item_sa = " & seq_item_sa
		set objRs = server.CreateObject("ADODB.RECORDSET")
		objRs.Open strSql,cn,0,1
		if not objRs.EOF then
			descr_item_sa						= objRs.Fields("descr_item_sa").Value
		end if
		objRs.Close
		set objRs = nothing
		'Escrever campos desabilitados e campos hidden 
		%>
    <tr> 
			<td class="texto" width="20%">Número do Processo:
			</td>
			<td class="texto" width="80%"><%=NumProcesso%></td>
    </tr>
    <tr> 
			<td class="texto" width="20%">Ítem da SA:
			</td>
			<td class="texto" width="80%"><%=descr_item_sa%></td>
    </tr>
    <tr> 
			<td class="texto" width="20%">Comentário:
			</td>
			<td class="texto" width="80%"><%=descr_Comentario%></td>
    </tr>
    
    <tr><td class="texto">&nbsp;
		<input type="hidden" name="Acao" value="<%=Acao%>">
		<input type="hidden" name="NumProcesso" value="<%=NumProcesso%>">
		<input type="hidden" name="sig_orgao_processo" value="<%=sig_orgao_processo%>">
		<input type="hidden" name="seq_Processo" value="<%=seq_Processo%>">
		<input type="hidden" name="ano_processo" value="<%=ano_processo%>">
		<input type="hidden" name="seq_sa" value="<%=seq_sa%>">
		<input type="hidden" name="seq_sa_complementar" value="<%=seq_sa_complementar%>">
		<input type="hidden" name="seq_area" value="<%=seq_area%>">
		<input type="hidden" name="seq_assunto" value="<%=seq_assunto%>">
		<input type="hidden" name="seq_item_sa" value="<%=seq_item_sa%>">
		<input type="hidden" name="seq_checklist" value="<%=seq_checklist%>">
		<input type="hidden" name="seq_Comentario" value="<%=seq_Comentario%>">
		<input type="hidden" name="txtComentario" value="<%=descr_Comentario%>">
		<input type="hidden" name="txtAnexo" value="<%=descr_anexo%>">
		<input type="hidden" name="txtRecomendacao" value="<%=descr_Recomendacao%>">
		<input type="hidden" name="chk_relatorio" value="<%=ind_relatorio%>">
		<input type="hidden" name="chk_ata" value="<%=ind_ata%>">
		<input type="hidden" name="Origem" value="<%=Origem%>">
		<input type="hidden" name="retorno" value="<%=retorno%>">
		<input type="hidden" name="ListaItensSa" value="<%=ListaItensSa%>">
		
    </td></tr>
		<%
	case "sa_item_auditoria_listar"
		'Tabela de sa_item_auditoria
		Acao											= Request("Acao")
		sig_orgao_processo				= Request("Sig_Orgao_Processo")
		seq_Processo							= Request("Seq_Processo")
		ano_processo							= Request("Ano_Processo")
		seq_sa										= Request("Seq_Sa")
		seq_sa_complementar				= Request("seq_sa_complementar")
		seq_area									= Request("SelArea")
		seq_assunto								= Request("SelAssunto")
		ListaItens								= Request("hidListaItens")
		BotaoAux									= Request("BotaoAux")
		NumProcesso = "PA" & "-" & sig_orgao_processo & "-" & right("00" & seq_Processo,3) & "/" & ano_processo
		'Escrever campos desabilitados e campos hidden 
		%>
    <tr> 
			<td class="texto" width="20%">Número do Processo:
			</td>
			<td class="texto" width="80%"><%=NumProcesso%></td>
    </tr>
    <tr><td class="texto">&nbsp;
		<input type="hidden" name="Acao" value="<%=Acao%>">
		<input type="hidden" name="NumProcesso" value="<%=NumProcesso%>">
		<input type="hidden" name="Sig_Orgao_Processo" value="<%=sig_orgao_processo%>">
		<input type="hidden" name="Seq_Processo" value="<%=seq_Processo%>">
		<input type="hidden" name="Ano_Processo" value="<%=ano_processo%>">
		<input type="hidden" name="Seq_Sa" value="<%=seq_sa%>">
		<input type="hidden" name="seq_sa_complementar" value="<%=seq_sa_complementar%>">
		<input type="hidden" name="SelArea" value="<%=seq_area%>">
		<input type="hidden" name="SelAssunto" value="<%=seq_assunto%>">
		<input type="hidden" name="hidListaItens" value="<%=ListaItens%>">
		<input type="hidden" name="BotaoAux" value="<%=BotaoAux%>">
		
    </td></tr>
		<%
	case "papel_trabalho_listar"
		'Tabela de sa_item_auditoria
		seq_sa_complementar								= Request("seq_sa_complementar")
		seq_assunto												= Request("seq_assunto")
		Origem														= Request("Origem")
		
		Sig_Orgao_Processo								= Request("Sig_Orgao_Processo")
		Ano_Processo											= Request("Ano_Processo")
		Seq_Processo											= Request("Seq_Processo")
		Seq_Sa														= Request("Seq_Sa")
		hidSeqItemSa											= Request("hidSeqItemSa")
		hidSeqArea												= Request("hidSeqArea")
		hidSeqAssunto											= Request("hidSeqAssunto")
		hidSeqChecklist										= Request("hidSeqChecklist")
		NumProcesso = "PA" & "-" & Sig_Orgao_Processo & "-" & right("00" & seq_Processo,3) & "/" & ano_processo
		'Escrever campos desabilitados e campos hidden 
		%>
    <tr> 
			<td class="texto" width="20%">Número do Processo:
			</td>
			<td class="texto" width="80%"><%=NumProcesso%></td>
    </tr>
    <tr><td class="texto">&nbsp;
    
		<input type="hidden" name="Sig_Orgao_Processo" value="<%=Sig_Orgao_Processo%>">
		<input type="hidden" name="Ano_Processo" value="<%=Ano_Processo%>">
		<input type="hidden" name="Seq_Processo" value="<%=Seq_Processo%>">
		<input type="hidden" name="Seq_Sa" value="<%=Seq_Sa%>">
		<input type="hidden" name="SelSa" value="<%=Seq_Sa%>">
		<input type="hidden" name="hidSeqItemSa" value="<%=hidSeqItemSa%>">
		<input type="hidden" name="hidSeqArea" value="<%=hidSeqArea%>">
		<input type="hidden" name="hidSeqAssunto" value="<%=hidSeqAssunto%>">
		<input type="hidden" name="hidSeqChecklist" value="<%=hidSeqChecklist%>">
		<input type="hidden" name="NumProcesso" value="<%=NumProcesso%>">

		<input type="hidden" name="seq_sa_complementar" value="<%=seq_sa_complementar%>">
		<input type="hidden" name="seq_assunto" value="<%=seq_assunto%>">
		<input type="hidden" name="Origem" value="<%=Origem%>">
		<input type="hidden" name="SelArea" value="<%=hidSeqArea%>">
		<input type="hidden" name="SelAssunto" value="<%=hidSeqAssunto%>">
		<input type="hidden" name="Seq_Area_Aux" value="<%=hidSeqArea%>">
		<input type="hidden" name="Seq_Assunto_Aux" value="<%=hidSeqAssunto%>">

    </td></tr>
		<%
	case "comentario_item_sa_listar"
		'Tabela de comentario_item_sa_listar
		Aux										= Request("Aux")
		NumProcesso						= Request("NumProcesso")
		sig_orgao_processo		= Request("sig_orgao_processo")
		seq_Processo					= Request("seq_Processo")
		ano_processo					= Request("ano_processo")
		seq_sa								= Request("seq_sa")
		seq_sa_complementar		= Request("seq_sa_complementar")
		seq_area							= Request("seq_area")
		seq_assunto						= Request("seq_assunto")
		seq_item_sa						= Request("seq_item_sa")
		seq_checklist					= Request("seq_checklist")
		seq_Comentario				= Request("seq_Comentario")
		Origem								= Request("Origem")
		retorno								= Request("retorno")
		ListaItensSa					= Request("ListaItensSa")
		
		'Escrever campos desabilitados e campos hidden 
		%>
    <tr> 
			<td class="texto" width="20%">Número do Processo:
			</td>
			<td class="texto" width="80%"><%=NumProcesso%></td>
    </tr>
    <tr><td class="texto">&nbsp;
    
		<input type="hidden" name="Aux" value="<%=Aux%>">
		<input type="hidden" name="NumProcesso" value="<%=NumProcesso%>">
		<input type="hidden" name="sig_orgao_processo" value="<%=sig_orgao_processo%>">
		<input type="hidden" name="seq_Processo" value="<%=seq_Processo%>">
		<input type="hidden" name="ano_processo" value="<%=ano_processo%>">
		<input type="hidden" name="seq_sa" value="<%=seq_sa%>">
		<input type="hidden" name="seq_sa_complementar" value="<%=seq_sa_complementar%>">
		<input type="hidden" name="seq_area" value="<%=seq_area%>">
		<input type="hidden" name="seq_assunto" value="<%=seq_assunto%>">
		<input type="hidden" name="seq_item_sa" value="<%=seq_item_sa%>">
		<input type="hidden" name="seq_checklist" value="<%=seq_checklist%>">
		<input type="hidden" name="seq_Comentario" value="<%=seq_Comentario%>">
		<input type="hidden" name="Origem" value="<%=Origem%>">
		<input type="hidden" name="retorno" value="<%=retorno%>">
		<input type="hidden" name="ListaItensSa" value="<%=ListaItensSa%>">
		
    </td></tr>
		<%
		
	end select

%>
