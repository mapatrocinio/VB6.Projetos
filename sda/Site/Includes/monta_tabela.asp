<%
	'ind_tabela assume de qual página veio para montar a tabela.
	'Assume os valores :
	'	ind_tabela = "pergunta"
	'	ind_tabela = "paaai"
	'	ind_tabela = "sa_item_auditoria"
	'	ind_tabela = "comentario_item_sa"
	'	ind_tabela = "relatorio_auditoria"
	'	ind_tabela = "template"
	dim strSql
	dim objRs
	dim controle
	controle = request("controle")
	if controle & "" = "E" then
	
	end if
	select case ind_tabela
	case "template_relatorio"
		'Tabela de Template relatório
		strSql = "select planilha.* from planilha_template_relatorio inner join planilha " & _
			"on planilha.seq_planilha = planilha_template_relatorio.seq_planilha " & _
			"where seq_tipo_documento = " & seq_tipo_documento & _
			" and seq_topico_doc = " & seq_topico_doc_conclu & _
			" and seq_template = " & seq_template & _
			" and ano_processo = " & ano_processo & _
			" and sig_orgao_processo = " & sig_orgao_processo & _
			" and seq_Processo = " & seq_Processo & _
			" order by planilha_template_relatorio.num_ordem" 

	case "template"
		'Tabela de Template
		strSql = "select planilha.* from planilha_template_doc_auditoria inner join planilha " & _
			"on planilha.seq_planilha = planilha_template_doc_auditoria.seq_planilha " & _
			"where seq_tipo_documento = " & str_seq_tipo_documento & _
			" and seq_topico_doc = " & str_seq_topico_documento & _
			" and seq_template = " & str_hid_seq_template & _
			" order by planilha_template_doc_auditoria.num_ordem" 
	case "relatorio_auditoria"
		'Tabela de relatorio_auditoria
		strSql = "select planilha.* from planilha_relatorio_auditoria inner join planilha " & _
			"on planilha.seq_planilha = planilha_relatorio_auditoria.seq_planilha " & _
			"where ano_processo = " & ano_processo & _
			"and sig_orgao_processo = '" & sig_orgao_processo & "'" & _
			"and seq_processo = " & seq_processo & _
			"and seq_relatorio = " & seq_relatorio & _
			" order by planilha_relatorio_auditoria.num_ordem" 
	case "pergunta"
		'Tabela de pergunta
		strSql = "select planilha.* from planilha_item_sa inner join planilha " & _
			"on planilha.seq_planilha = planilha_item_sa.seq_planilha " & _
			"where seq_assunto = " & seq_Assunto & _
			"and seq_item_sa = " & seq_item_sa & _
			" order by planilha_item_sa.num_ordem" 
	case "paaai"
		'Tabela de paaai
		strSql = "select planilha.* from planilha_paaai inner join planilha " & _
			"on planilha.seq_planilha = planilha_paaai.seq_planilha " & _
			"where ano_paaai = " & ano_busca & _
			"and cod_uo = '" & unidade_busca & _
			"' order by planilha_paaai.num_ordem" 
	case "sa_item_auditoria"
		'Tabela de sa_item_auditoria
		strSql = "select planilha.* from planilha_sa_item_auditoria inner join planilha " & _
			"on planilha.seq_planilha = planilha_sa_item_auditoria.seq_planilha " & _
			"where planilha_sa_item_auditoria.ano_processo = " & ano_processo & _
			" and planilha_sa_item_auditoria.sig_orgao_processo = '" & sig_orgao_processo & "'" & _
			" and planilha_sa_item_auditoria.seq_processo = " & seq_Processo  & _
			" and planilha_sa_item_auditoria.seq_sa = " & seq_sa  & _
			" and planilha_sa_item_auditoria.seq_sa_complementar = '" & seq_sa_complementar & "'" & _
			" and planilha_sa_item_auditoria.seq_assunto = " & seq_assunto  & _
			" and planilha_sa_item_auditoria.seq_item_sa = " & seq_item_sa  & _
			" and planilha_sa_item_auditoria.seq_area = " & seq_area & _
			" order by planilha_sa_item_auditoria.num_ordem" 
	case "comentario_item_sa"
		'Tabela de comentario_item_sa
		strSql = "select planilha.* from planilha_comentario_item_sa inner join planilha " & _
			"on planilha.seq_planilha = planilha_comentario_item_sa.seq_planilha " & _
			"where planilha_comentario_item_sa.ano_processo = " & ano_processo & _
			" and planilha_comentario_item_sa.sig_orgao_processo = '" & sig_orgao_processo & "'" & _
			" and planilha_comentario_item_sa.seq_processo = " & seq_Processo  & _
			" and planilha_comentario_item_sa.seq_sa = " & seq_sa  & _
			" and planilha_comentario_item_sa.seq_sa_complementar = '" & seq_sa_complementar & "'" & _
			" and planilha_comentario_item_sa.seq_assunto = " & seq_assunto  & _
			" and planilha_comentario_item_sa.seq_item_sa = " & seq_item_sa  & _
			" and planilha_comentario_item_sa.seq_comentario = " & seq_comentario  & _
			" and planilha_comentario_item_sa.seq_area = " & seq_area & _
			" order by planilha_comentario_item_sa.num_ordem" 
			
	end select
	set objRs = server.CreateObject("ADODB.RECORDSET")
	objRs.Open strSql,cn,0,1%>
	<TABLE border=0 cellPadding=1 cellSpacing=1 width="100%">
		<TR>
			<TD WIDTH=40%><table border=0 cellPadding=1 cellSpacing=1 width="100%"><tr><td background="imagens/azul_fundo1.gif" style="HEIGHT: 1px" height=1></td></tr></table></TD>
			<TD nowrap class="texto" align="center" style="WIDTH: 20%"><b>Tabelas Associadas</b></TD>
			<TD WIDTH=40%><table border=0 cellPadding=1 cellSpacing=1 width="100%"><tr><td background="imagens/azul_fundo1.gif" style="HEIGHT: 1px" height=1></td></tr></table></TD>
		</TR>
	</TABLE>
	<%
	if not objRs.EOF then%>
		<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor=#f0f0f0>
		<%
		do while not objRs.EOF
			%>
			<tr onmouseover="Mouseover(this,'');" onmouseout="Mouseout(this);">
				<td onclick="javascript:navegar_tabela('A',<%=objRs.Fields("seq_planilha")%>);" style="cursor: Hand" width="30px" class="texto" align="center" valign="middle"><img src="imagens/icone_alterar.jpg" alt="Alterar tabela" border=0></td>
				<td onclick="javascript:Tabelas(<%=objRs.Fields("seq_planilha")%>);" style="cursor: Hand" width="30px" class="texto" align="center" valign="middle"><img src="imagens/icone_visualizar.jpg" alt="Vizualizar tabelas" border=0></td>
				<td onclick="javascript:navegar_tabela('EXCLUIR',<%=objRs.Fields("seq_planilha")%>);" style="cursor: Hand" width="30px" class="texto" align="center" valign="middle"><img src="imagens/icone_excluir.jpg" alt="Excluir tabela" border=0></td>
				<td onclick="javascript:navegar_tabela('A',<%=objRs.Fields("seq_planilha")%>);" style="cursor: Hand" class="texto" valign="middle">&nbsp;<%=objRs.Fields("descr_titulo")%></td>
			</tr>
		<%
			objRs.MoveNext
		loop
		%>
		</table>
		<%
		objRs.Close
		set objRs = nothing
	end if
		%>
	<br>
	<%
	botao "", "", "botaoazul", "javascript:navegar_tabela('I','');", "Nova Tabela", "nova_tabela",""
	%>
	<SCRIPT LANGUAGE=javascript>
<!--
function Tabelas(seq_planilha){
	navegar_tabela("L", seq_planilha);
	//var form=document.forms[0];
	//form.action="detalhes_tabela_exibir.asp?ind_tabela=geral_listar&status_detalhes=inicial&acao_detalhes=L&seq_planilha="+seq_planilha;
	//form.submit();
}

function navegar_tabela(acao, seq_planilha){
	var form = document.forms[0];
	<%if ind_tabela = "sa_item_auditoria"	then%>
		if (blnMudouTexto==true){
			if (confirm("ATENÇÃO: O texto do ítem foi modificado. Caso você prossiga o texto digitado será perdido.\n\nDeseja prosseguir sem Salvar as alterações?")==false){
				//não confirmou seguir sem salvar alterações
				return;
			}
		}
	<%end if%>
	if (acao=='L'){
		form.action="detalhes_tabela_exibir.asp?ind_tabela=geral_listar&status_detalhes=inicial&acao_detalhes=L&seq_planilha="+seq_planilha;
		form.submit();
	}else	if (acao=='EXCLUIR'){
//		alert("<%=ind_tabela%>");
		<%
		select case ind_tabela
		case "template_relatorio"
			'Tabela de template_relatório
		%>
			if (confirm("Tem certeza que deseja excluir esta tabela?")) {
				form.action = "RelatorioSalvar.asp?seq_planilha="+seq_planilha+"&acao_tabela=ExcluirTabela";
				form.submit();
			}
		<%
		case "template"
			'Tabela de template
		%>
			if (confirm("Tem certeza que deseja excluir esta tabela?")) {
				form.controle_template_cadastrar.value="EXCLUIR_TABELA"
				form.action = "template_cadastrar.asp?seq_planilha="+seq_planilha+"&acao_tabela=ExcluirTabela";
				form.submit();
			}
		<%
		case "relatorio_auditoria"
			'Tabela de relatorio_auditoria
		%>
			if (confirm("Tem certeza que deseja excluir esta tabela?")) {
				form.action = "ItemRelatorioSalvar.asp?seq_planilha="+seq_planilha+"&acao_tabela=ExcluirTabela";
				form.submit();
			}
		<%
		case "pergunta"
			'Tabela de pergunta 
		%>
			if (confirm("Tem certeza que deseja excluir esta tabela?")) {
				form.action = "AdmPergSalvar.asp?seq_planilha="+seq_planilha+"&acao_tabela=ExcluirTabela";
				form.submit();
			}
		<%
		case "paaai"
			'Tabela de paaai
		%>
			if (confirm("Tem certeza que deseja excluir esta tabela?")) {
				form.action = "AdmPaaaiSalvar.asp?seq_planilha="+seq_planilha+"&acao_tabela=ExcluirTabela";
				form.submit();
			}
		<%
		case "sa_item_auditoria"
			'Tabela de sa_item_auditoria
		%>
			if (confirm("Tem certeza que deseja excluir esta tabela?")) {
				form.action = "SASalvar3.asp?seq_planilha="+seq_planilha+"&acao_tabela=ExcluirTabela";
				form.submit();
			}
		<%
		case "comentario_item_sa"
			''Tabela de comentario_item_sa
		%>
			if (confirm("Tem certeza que deseja excluir esta tabela?")) {
				form.action = "ComentarioSalvar.asp?seq_planilha="+seq_planilha+"&acao_tabela=ExcluirTabela";
				form.submit();
			}
		<%
		end select
		%>
	}
	else {
		<%
		select case ind_tabela
		case "template"
		%>
			form.controle_template_cadastrar.value = "INSERIR_TABELA";
		<%
		end select%>
		form.action = "dados_tabela.asp?acao_dados="+acao+"&status_dados=inicial&seq_planilha="+seq_planilha
		form.submit();
	}
}
function Mouseout(linha){
	linha.bgColor = "#ffffff";
}

function Mouseover(linha,tipotabela){
	if (tipotabela == "") 
		linha.bgColor = "#e0ffff"; 
	else 
		linha.bgColor = "#e9e1f0"; 
}

//-->
</SCRIPT>

