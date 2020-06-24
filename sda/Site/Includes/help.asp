<!--#include file="../includes/funcoes.asp"-->
<%
function help(pagina)
	Dim strRetorno
	Dim PulaLinha
	PulaLinha = "<BR>"
	strRetorno = "" & vbcrlf 
	select case pagina
	case "area_incluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar nova área na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Areas Cadastradas
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Áreas Cadastradas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Serve somente para consulta das áreas que já foram cadastradas.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Descrição da Área
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Descrição da Área</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar a descrição da área.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"

	case "area_alterar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar a área na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Areas Cadastradas
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Areas Cadastradas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção, a área que deseja alterar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Descrição da Área
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Descrição da Área</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar a descrição da área.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
	case "area_excluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-1-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, excluir a área na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Areas 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help_Help' href=#1 name=#1>Áreas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção, a área que deseja excluir.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"

	case "area_encaminhamento_incluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar a área para encaminhamento na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Areas Cadastradas
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Áreas Cadastradas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Serve somente para consulta das áreas para encaminhamento que já foram cadastradas.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Descrição da Área
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Descrição da Área</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar a descrição da área para encaminhamento.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"

	case "area_encaminhamento_alterar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar a área para encaminhamento na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Areas Cadastradas
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Áreas Cadastradas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção, para selecionar a área para encaminhamento que deseja alterar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Descrição da Área
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Descrição da Área</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar a descrição da área.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
	case "area_encaminhamento_excluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-1-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, excluir a área para encaminhamento na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Areas 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help_Help' href=#1 name=#1>Áreas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção, a área para encaminhamento que deseja excluir.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"

	case "area_encaminhamento_consultar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, consultar a área para  para encaminhamento para poder incluir, alterar ou excluir.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Areas 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Áreas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção, a área para encaminhamento que deseja consultar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Descrição da Área
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Descrição da Área</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar a descrição para incluir ou alterar a área.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "assunto_incluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar o assunto na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Serve somente para consulta dos assuntos que já foram cadastrados.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Descrição do Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Descrição do Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar a descrição do assunto.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"

	case "assunto_alterar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar o assunto na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção, o assunto que deseja alterar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Descrição do Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Descrição do Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar a descrição do assunto.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
	case "assunto_excluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-1-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, excluir o assunto na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção, o assunto que deseja excluir.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "checklist_incluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar ítens de checklist na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar o assunto para o qual gostaria de cadastrar um ítem de checklist.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Checklist Cadastrados
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Checklist Cadastrados</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Serve somente para consulta dos ítens de checklist que já foram cadastrados.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Descrição do ítem
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Descrição do ítem</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar a descrição do ítem do checklist.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"

	case "checklist_alterar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar o ítem de checklist na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção, o assunto para filtrar os ítens de checklist.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Checklist Cadastrados
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Checklist Cadastrados</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção, o ítem de  checklist cadastrado para alteração.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Descrição do ítem
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Descrição do ítem</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar a descrição do ítem do checklist.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
	case "checklist_excluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-1-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, excluir checklist na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção, o assunto para filtrar os ítens de checklist.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Checklist Cadastrados
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Checklist Cadastrados</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção, o ítem de checklist que deseja excluir.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
	
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "paaai_incluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar PAAAI (Plano Anual de Atividades de Auditoria Interna) na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Mês de Execução
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Mês de Execução</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar o mês de execução.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Ano do PAAAI
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Ano do PAAAI</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar o ano do PAAAI.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data de Aprovação
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data de Aprovação</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar a data de aprovação.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Unidade
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Unidade</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar o botão de seleção RNML ou UP, para o sistema exibir na caixa de seleção as unidades existentes. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Local de Trabalho
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Local de Trabalho</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar o local de trabalho.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Observação
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#6 name=#6>Observação</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar alguma observação.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"

	case "paaai_alterar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar PAAAI na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Ano para busca 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Ano para busca </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar o ano que deseja para que sejam listados todos os PAAAI’s para o mesmo. Depois disso, clicar na imagem ao lado do campo, para efetuar a busca.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Unidade
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Unidade</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar o botão de seleção da unidade para fazer a busca.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Mês de Execução
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Mês de Execução</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar o mês de execução.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Ano do PAAAI
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Ano do PAAAI</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar o ano do PAAAI.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data de Aprovação
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Data de Aprovação</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar a data de aprovação.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Unidade
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#6 name=#6>Unidade</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar o botão de seleção RNML ou UP, para o sistema exibir na caixa de seleção as unidades para seleção referente ao botão selecionado. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Local de Trabalho
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#7 name=#7>Local de Trabalho</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar o local de trabalho.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Observação
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#8 name=#8>Observação</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar alguma observação.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
	case "paaai_excluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-1-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, excluir PAAAI na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Ano para a Busca
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Ano para a Busca</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar o ano que deseja para que sejam listados todos os PAAAI’s para o mesmo. Depois disso, clicar na imagem ao lado do campo, para efetuar a busca.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Unidade
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Unidade</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar o botão de seleção da unidade para fazer a busca.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Mês de Execução
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Mês de Execução</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o mês de execução cadastrado para esse PAAAI.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Ano do PAAAI
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Ano do PAAAI</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe ano cadastrado para esse PAAAI.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data de Aprovação
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Data de Aprovação</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe a data de aprovação cadastrada para esse PAAAI.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Unidade
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#6 name=#6>Unidade</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe a unidade cadastrada para esse PAAAI. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Local de Trabalho
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#7 name=#7>Local de Trabalho</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o local de trabalho cadastrado para esse PAAAI.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Observação
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#8 name=#8>Observação</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe a observação cadastrada para esse PAAAI.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
	
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"

	case "paaai_consultar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-1-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, consultar o PAAAI para poder incluir, alterar ou excluir.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Ano para a Busca
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Ano para a Busca</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar o ano que deseja para que sejam listados todos os PAAAI’s para o mesmo. Depois disso, clicar na imagem ao lado do campo, para efetuar a busca.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Unidade
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Unidade</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar o botão de seleção a unidade para fazer a busca.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Mês de Execução
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Mês de Execução</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o mês de execução cadastrado para esse PAAAI.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Ano do PAAAI
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Ano do PAAAI</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe ano cadastrado para esse PAAAI.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data de Aprovação
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Data de Aprovação</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe a data de aprovação cadastrada para esse PAAAI.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Unidade
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#6 name=#6>Unidade</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe a unidade cadastrada para esse PAAAI. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Local de Trabalho
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#7 name=#7>Local de Trabalho</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o local de trabalho cadastrado para esse PAAAI.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Observação
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#8 name=#8>Observação</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe a observação cadastrada para esse PAAAI.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 	
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "perguntar_incluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar perguntas na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção o assunto para o qual deseja cadastrar uma pergunta.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Perguntas Cadastradas
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Perguntas Cadastradas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Serve somente para consulta das perguntas que já foram cadastradas.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Descrição da Pergunta
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Descrição da Pergunta</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar a descrição da pergunta.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
				
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"

	case "perguntar_alterar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar as perguntas cadastradas na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção o assunto.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Perguntas Cadastradas
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Perguntas Cadastradas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode selecionar na caixa de seleção, uma pergunta cadastrada para alteração.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Descrição da Pergunta
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Descrição da Pergunta</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar a descrição da pergunta.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
	case "perguntar_excluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-1-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, excluir pergunta cadastrada na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção o assunto.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Perguntas Cadastradas
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Perguntas Cadastradas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode selecionar na caixa de seleção, uma pergunta cadastrada para exclusão.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Descrição da Pergunta
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Descrição da Pergunta</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe a descrição da pergunta cadastrada.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "perguntar_checklist_incluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar perguntas X ítem de checklist na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção o assunto.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Pergunta 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Pergunta</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode selecionar na caixa de seleção uma pergunta cadastrada.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"

	case "perguntar_checklist_consultar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, consultar perguntas X ítem de checklist na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode consultar selecionando na caixa de seleção o assunto.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Pergunta 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Pergunta</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode consultar selecionando na caixa de seleção uma pergunta cadastrada.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
	case "perguntar_checklist_consultar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-1-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, incluir ou alterar a pergunta x checklist.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Selecione uma empresa na caixa de seleção, caso o usuário queira ativar uma pessoa física.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Pergunta 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Pergunta</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Selecione uma empresa na caixa de seleção, caso o usuário queira ativar uma pessoa física.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
	
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "template_incluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar o template na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Tipo de Template
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Tipo de Template</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção o tipo de template. O Tipo de template pode ser: SA , SA Complementar, Papel de Trabalho e  Relatório de Auditoria.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Tópico do Documento 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Tópico do Documento</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode selecionar na caixa de seleção um tópico do documento. O tópico do documento pode ser por exemplo, a base legal dentro da SA. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Campos Chave
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Campos Chave</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção o campo chave. Caso queira adicionar no texto no campo descrição a palavra chave selecionado, basta clicar no Link adicionar ao texto. O campo chave serve para que o sistema pegue automaticamente da base de dados, informações desejadas sem que o usuário precise digitá-las.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Templates Cadastrados
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Templates Cadastrados</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Serve apenas para consultar os templates já cadastrados na base de dados.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Descrição
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Descrição</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar o conteúdo do template.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
				
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"

	case "template_alterar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar o template na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Tipo de Template
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Tipo de Template</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção o tipo de template.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Tópico do Documento 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Tópico do Documento</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode selecionar na caixa de seleção um tópico do documento.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Campos Chave
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Campos Chave</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção o campo chave. Caso queira adicionar no texto no campo descrição a palavra chave selecionado, basta clicar no Link adicionar ao texto.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Templates Cadastrados
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Templates Cadastrados/a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção o template cadastrados.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Descrição
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Descrição</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar alguma descrição se desejar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
	case "template_excluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-1-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, excluir o template na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Tipo de Template
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Tipo de Template</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção o tipo de template.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Tópico do Documento 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Tópico do Documento</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode selecionar na caixa de seleção um tópico do documento.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Templates Cadastrados
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Templates Cadastrados</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção o template que deseja alterar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Descrição
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Descrição</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar alguma descrição se desejar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"

	case "processo_Desbloquear"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, desbloquear o processo na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na imagem ao lado do campo o número do processo que queira desbloquear.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "exportar_processo"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, exportar processo da base de dados para o módulo notebook.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado para exportar clicando na imagem ao lado do campo. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "importar_processo"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, importar o processo, depois de unificado dos notebooks para a base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Importar Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Importar Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve clicar no Botão Procurar para escolher o arquivo (auditoria.mdb que foi unificado após a auditoria)  que queira importar. Após isto, o sistema exibe essa tela, para escolher o arquivo para importar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "processo_incluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar os processos na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Data de Inicio
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Data de Inicio</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar a data inicio do processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Data Fim
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Data Fim</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar a data fim do processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
	
		'Unidade Auditada
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Unidade Auditada</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar o botão de seleção RNML ou UP, para o sistema exibir na caixa de seleção as unidades existentes. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Tipo de Auditoria
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Tipo de Auditoria</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar o botão de seleção ordinária ou extraordinária, para marcar o tipo de Auditoria.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Auditor Responsável
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Auditor Responsável</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar o auditor responsável.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Função
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#6 name=#6>Função</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar a função do auditor.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
				
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"

	case "processo_alterar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar o processo cadastrado na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela para alterar ou bloquear o processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data de Inicio
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Data de Inicio</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar a data inicio do processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Data Fim
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data Fim</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar a data fim do processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
	
		'Unidade Auditada
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Unidade Auditada</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode selecionar o botão de seleção RNML ou UP, para o sistema exibir na caixa de seleção as unidades auditadas para seleção referente ao botão selecionado. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Tipo de Auditoria
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Tipo de Auditoria</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode selecionar o botão de seleção ordinária ou extraordinária, para marcar o tipo de Auditoria.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Auditor Responsável
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#6 name=#6>Auditor Responsável</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar o auditor responsável.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Função
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#7 name=#7>Função</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar a função do auditor.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "processo_excluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, excluir o processo cadastrado na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
			
		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela para excluir o processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data de Inicio
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Data de Inicio</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe a data inicio cadastrada para o processo selecionado.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Data Fim
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data Fim</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe a data fim cadastrada para o processo selecionado.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
	
		'Unidade Auditada
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Unidade Auditada</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe a unidade auditada cadastrada para o processo selecionado. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Tipo de Auditoria
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Tipo de Auditoria</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o tipo de auditoria cadastrado para o processo selecionado.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Auditor Responsável
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#6 name=#6>Auditor Responsável</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário exibe o auditor responsável.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Função
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#7 name=#7>Função</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário exibe a função do auditor.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "processo_consultar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, consultar o processo.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Selecione uma empresa na caixa de seleção, caso o usuário queira ativar uma pessoa física.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data de Inicio
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Data de Inicio</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Selecione uma empresa na caixa de seleção, caso o usuário queira ativar uma pessoa física.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Data Fim
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data Fim</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Selecione uma empresa na caixa de seleção, caso o usuário queira ativar uma pessoa física.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
	
		'Unidade Auditada
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Unidade Auditada</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Selecione uma empresa na caixa de seleção, caso o usuário queira ativar uma pessoa física.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Tipo de Auditoria
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Tipo de Auditoria</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Selecione uma empresa na caixa de seleção, caso o usuário queira ativar uma pessoa física.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Auditor Responsável
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#6 name=#6>Auditor Responsável</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Auditor responsável cadastrada  para o processo selecionado.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Função
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#7 name=#7>Função</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Função cadastrada  para o processo selecionado.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "processo_membro_buscar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, pesquisar pela unidade os auditores e selecionar os membros da equipe referente ao processo.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Unidade 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Unidade</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode selecionar na caixa de seleção a unidade desejada para pesquisa.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Auditor
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Auditor</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar o nome do auditor para pesquisar ou clicar na imagem ao lado do campo para pesquisar o auditor.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "info_auditor_salvar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar as informações do auditor</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Conselho Regional 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Conselho Regional</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar o conselho regional do auditor.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Número do Registro 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Número do Registro</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar o número do registro do auditor.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'UF 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>UF</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar na caixa de seleção a unidade federativa do auditor.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
	
		'Função 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Função</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar a função do auditor.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "processo_buscar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, pesquisar o processo.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela para excluir o processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data de Inicio
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Data de Inicio</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe a data inicio cadastrada para o processo selecionado.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Data Fim
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data Fim</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe a data fim cadastrada para o processo selecionado.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
	
		'Unidade Auditada
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Unidade Auditada</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe a unidade auditada cadastrada para o processo selecionado. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Tipo de Auditoria
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Tipo de Auditoria</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o tipo de auditoria cadastrado para o processo selecionado.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "sa_incluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar a SA na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
			
		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela para excluir o processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Número da SA
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Número da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Selecione uma empresa na caixa de seleção, caso o usuário queira ativar uma pessoa física.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode cadastrar a data da SA.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Tipo da SA
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Tipo da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode selecionar o botão de seleção normal ou complementar, para marcar o tipo de SA.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
				
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"

	case "sa_alterar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar os dados da SA na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela para alterar o processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Número da SA
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Número da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar a SA que deseja alterar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar a data da SA.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Tipo da SA
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Tipo da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode selecionar o botão de seleção normal ou complementar, para marcar o tipo de SA.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "sa_excluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, excluir a SA na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela para excluir o processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Número da SA
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Número da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode selecionar qual SA ele deseja excluir.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Data cadastrada para a SA selecionada.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Tipo da SA
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Tipo da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Tipo de SA cadastrado para a SA selecionada.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "sa_consultar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, consultar ou imprimir a SA cadastrada na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
			
		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela para excluir o processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Número da SA
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Número da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode selecionar qual SA ele deseja consultar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Data cadastrada para a SA selecionada.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Tipo da SA
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Tipo da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Tipo de SA cadastrado para a SA selecionada.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
			
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
			
	case "checklist_gerar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, gerar o checklist para o processo na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
			
		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Número da SA 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Número da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Após selecionar o número do processo, o usuário deve selecionar na caixa de seleção o número da SA.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Após selecionar a SA, o usuário deve selecionar na caixa de seleção o assunto.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Ítem
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Ítem</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Após selecionar os campos anteriores, o sistema exibe uma listagem dos Itens, referente aos campos selecionados.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
			
	case "checklist_consultar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, consultar o checklist na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
			
		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Número da SA 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Número da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Após selecionar o número do processo, o usuário deve selecionar na caixa de seleção o número da SA.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Após selecionar a SA, o usuário deve selecionar na caixa de seleção o assunto.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Ítem
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Ítem</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Após selecionar os campos anteriores, o sistema exibe uma listagem dos Itens, referente aos campos selecionados.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
				
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "papel_trabalho_listar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, exibir todos os ítens da SA disponíveis para o papel de trabalho.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
				
		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Número da SA 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Número da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Após selecionar o número do processo, o usuário deve selecionar na caixa de seleção o número da SA.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Área
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Área</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Após selecionar o número da SA, o usuário pode selecionar na caixa de seleção a área.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Após selecionar o número da SA, o usuário pode selecionar na caixa de seleção o assunto.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Ítens
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Ítens</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Após selecionar os campos, o sistema exibe uma listagem dos Itens da SA, referente aos campos selecionados.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
	
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
			
	case "comentario_incluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar o comentário na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o número do processo, referente ao item selecionado na tela anterior.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Ítem
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Ítem</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o ítem, referente ao item selecionado na tela anterior.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Ir para Relatório 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Ir para Relatório </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode selecionar opção se deseja que esse comentário vá para o relatório.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Ir para Ata de Auditoria 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Ir para Ata de Auditoria </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode selecionar na nesta caixa de texto se deseja que esse comentário vá para a Ata de Auditoria.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Comentário 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Comentário </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Para inserir um comentário, o usuário deve digitar neste campo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Recomendação  
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#6 name=#6>Recomendação  </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Para inserir uma recomendação, o usuário deve digitar neste campo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Anexos  
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#7 name=#7>Anexos </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Para informar se existem anexos existentes a esse comentário.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "comentario_alterar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar um comentário.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o número do processo, referente ao item selecionado na tela anterior.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Ítem
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Ítem</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o ítem, referente ao item selecionado na tela anterior.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Ir para Relatório 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Ir para Relatório </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode selecionar opção se deseja que esse comentário vá para o relatório.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Ir para Ata de Auditoria 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Ir para Ata de Auditoria </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode selecionar na nesta caixa de texto se deseja que esse comentário vá para a Ata de Auditoria.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Comentário 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Comentário </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Para inserir um comentário, o usuário deve digitar neste campo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Recomendação  
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#6 name=#6>Recomendação  </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Para inserir uma recomendação, o usuário deve digitar neste campo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Anexos  
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#7 name=#7>Anexos </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Para informar se existem anexos existentes a esse comentário.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "comentario_consultar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, consultar o comentário na base de dados. O sistema exibe os comentários referentes ao ítem selecionado na tela anterior</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o número do processo, referente ao item selecionado na tela anterior.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Número da SA
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Número da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o número da SA, referente ao item selecionado na tela anterior.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Item 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Item </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o  item, referente ao item selecionado na tela anterior.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "ata_incluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar o ítem da ata de Auditoria na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
				
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "ata_alterar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar o ítem da ata de Auditoria na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
				
		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
				
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "ata_excluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, excluir o ítem da ata de Auditoria na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
				
		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os ítens da ata de auditoria na página. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
				
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "ata_consultar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, consultar os ítens da ata ou imprimí-la.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
				
		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
				
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "relatorio_incluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, gerar um relatório.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela para preenchimento. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Número do Ofício
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Número do Ofício</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar o número do ofício.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Data do Ofício
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data do Ofício</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar a data do ofício.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "relatorio_alterar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar um relatório.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela para alterar. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Número do Ofício
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Número do Ofício</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar o número do ofício.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Data do Ofício
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data do Ofício</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar a data do ofício.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "relatorio_excluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, excluir um relatório.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
			
		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibirá as informações contidas no processo de auditoria  selecionado. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
				
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "relatorio_consultar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, consultar ou imprimir um relatório de auditoria.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela para preenchimento. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Número do Ofício
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Número do Ofício</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar o número do ofício para consultar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Data do Ofício
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data do Ofício</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar a data do ofício para consultar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
				
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"

	case "unificar_p_trabalho"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, unificar o processo.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Arquivo para Importação
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Arquivo para Importação</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Selecione uma empresa na caixa de seleção, caso o usuário queira ativar uma pessoa física.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
				
	case "parecer_incluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar o parecer da Auditoria na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data do Parecer
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Data do Parecer</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário deve digitar a data do parecer.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
				
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "parecer_alterar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar o parecer da Auditoria na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
			
		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data do Parecer
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Data do Parecer</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode alterar a data do parecer.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
				
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "parecer_excluir"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, excluir o parecer da Auditoria na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"
			
		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
	
	case "parecer_consultar"
		'Cabeçalho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, consultar o(s) parecer(es) da Auditoria na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descrição</td>"
		strRetorno = strRetorno & "</tr>"

		'Número do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Número do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usuário pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Após selecionado o número do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
	
		'Fechando o table	
		strRetorno = strRetorno & "<tr><td><br></td>"
	    strRetorno = strRetorno & "<td class='linha_clara' colspan=2>"
		strRetorno = strRetorno & "<div align='center'>"
		strRetorno = strRetorno & "<input class='amarelobotao' type='Button' Value='Fechar' class='botao' vspace='2' onclick='javascript:window.close();' id=Button1 name=Button1>"
		strRetorno = strRetorno & "</div>"
		strRetorno = strRetorno & "</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
			
	
	case else
		strRetorno = "Não há suporte para este campo."
	end select
	help = strRetorno
	
end function

%>