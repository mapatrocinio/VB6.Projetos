<!--#include file="../includes/funcoes.asp"-->
<%
function help(pagina)
	Dim strRetorno
	Dim PulaLinha
	PulaLinha = "<BR>"
	strRetorno = "" & vbcrlf 
	select case pagina
	case "area_incluir"
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar nova �rea na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Areas Cadastradas
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>�reas Cadastradas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Serve somente para consulta das �reas que j� foram cadastradas.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Descri��o da �rea
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Descri��o da �rea</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar a descri��o da �rea.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar a �rea na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Areas Cadastradas
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Areas Cadastradas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o, a �rea que deseja alterar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Descri��o da �rea
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Descri��o da �rea</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar a descri��o da �rea.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-1-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, excluir a �rea na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Areas 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help_Help' href=#1 name=#1>�reas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o, a �rea que deseja excluir.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar a �rea para encaminhamento na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Areas Cadastradas
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>�reas Cadastradas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Serve somente para consulta das �reas para encaminhamento que j� foram cadastradas.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Descri��o da �rea
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Descri��o da �rea</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar a descri��o da �rea para encaminhamento.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar a �rea para encaminhamento na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Areas Cadastradas
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>�reas Cadastradas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o, para selecionar a �rea para encaminhamento que deseja alterar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Descri��o da �rea
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Descri��o da �rea</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar a descri��o da �rea.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-1-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, excluir a �rea para encaminhamento na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Areas 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help_Help' href=#1 name=#1>�reas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o, a �rea para encaminhamento que deseja excluir.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, consultar a �rea para  para encaminhamento para poder incluir, alterar ou excluir.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Areas 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>�reas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o, a �rea para encaminhamento que deseja consultar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Descri��o da �rea
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Descri��o da �rea</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar a descri��o para incluir ou alterar a �rea.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Serve somente para consulta dos assuntos que j� foram cadastrados.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Descri��o do Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Descri��o do Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar a descri��o do assunto.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o, o assunto que deseja alterar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Descri��o do Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Descri��o do Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar a descri��o do assunto.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o, o assunto que deseja excluir.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar �tens de checklist na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar o assunto para o qual gostaria de cadastrar um �tem de checklist.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Checklist Cadastrados
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Checklist Cadastrados</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Serve somente para consulta dos �tens de checklist que j� foram cadastrados.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Descri��o do �tem
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Descri��o do �tem</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar a descri��o do �tem do checklist.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar o �tem de checklist na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o, o assunto para filtrar os �tens de checklist.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Checklist Cadastrados
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Checklist Cadastrados</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o, o �tem de  checklist cadastrado para altera��o.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Descri��o do �tem
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Descri��o do �tem</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar a descri��o do �tem do checklist.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o, o assunto para filtrar os �tens de checklist.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Checklist Cadastrados
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Checklist Cadastrados</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o, o �tem de checklist que deseja excluir.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'M�s de Execu��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>M�s de Execu��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar o m�s de execu��o.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Ano do PAAAI
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Ano do PAAAI</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar o ano do PAAAI.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data de Aprova��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data de Aprova��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar a data de aprova��o.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Unidade
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Unidade</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar o bot�o de sele��o RNML ou UP, para o sistema exibir na caixa de sele��o as unidades existentes. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Local de Trabalho
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Local de Trabalho</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar o local de trabalho.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Observa��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#6 name=#6>Observa��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar alguma observa��o.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Ano para busca 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Ano para busca </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar o ano que deseja para que sejam listados todos os PAAAI�s para o mesmo. Depois disso, clicar na imagem ao lado do campo, para efetuar a busca.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Unidade
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Unidade</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar o bot�o de sele��o da unidade para fazer a busca.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'M�s de Execu��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>M�s de Execu��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar o m�s de execu��o.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Ano do PAAAI
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Ano do PAAAI</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar o ano do PAAAI.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data de Aprova��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Data de Aprova��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar a data de aprova��o.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Unidade
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#6 name=#6>Unidade</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar o bot�o de sele��o RNML ou UP, para o sistema exibir na caixa de sele��o as unidades para sele��o referente ao bot�o selecionado. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Local de Trabalho
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#7 name=#7>Local de Trabalho</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar o local de trabalho.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Observa��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#8 name=#8>Observa��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar alguma observa��o.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Ano para a Busca
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Ano para a Busca</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar o ano que deseja para que sejam listados todos os PAAAI�s para o mesmo. Depois disso, clicar na imagem ao lado do campo, para efetuar a busca.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Unidade
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Unidade</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar o bot�o de sele��o da unidade para fazer a busca.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'M�s de Execu��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>M�s de Execu��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o m�s de execu��o cadastrado para esse PAAAI.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Ano do PAAAI
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Ano do PAAAI</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe ano cadastrado para esse PAAAI.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data de Aprova��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Data de Aprova��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe a data de aprova��o cadastrada para esse PAAAI.</td>" & vbcrlf 
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
		
		'Observa��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#8 name=#8>Observa��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe a observa��o cadastrada para esse PAAAI.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Ano para a Busca
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Ano para a Busca</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar o ano que deseja para que sejam listados todos os PAAAI�s para o mesmo. Depois disso, clicar na imagem ao lado do campo, para efetuar a busca.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Unidade
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Unidade</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar o bot�o de sele��o a unidade para fazer a busca.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'M�s de Execu��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>M�s de Execu��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o m�s de execu��o cadastrado para esse PAAAI.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Ano do PAAAI
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Ano do PAAAI</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe ano cadastrado para esse PAAAI.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data de Aprova��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Data de Aprova��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe a data de aprova��o cadastrada para esse PAAAI.</td>" & vbcrlf 
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
		
		'Observa��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#8 name=#8>Observa��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe a observa��o cadastrada para esse PAAAI.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o o assunto para o qual deseja cadastrar uma pergunta.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Perguntas Cadastradas
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Perguntas Cadastradas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Serve somente para consulta das perguntas que j� foram cadastradas.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Descri��o da Pergunta
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Descri��o da Pergunta</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar a descri��o da pergunta.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o o assunto.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Perguntas Cadastradas
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Perguntas Cadastradas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode selecionar na caixa de sele��o, uma pergunta cadastrada para altera��o.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Descri��o da Pergunta
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Descri��o da Pergunta</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar a descri��o da pergunta.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o o assunto.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Perguntas Cadastradas
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Perguntas Cadastradas</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode selecionar na caixa de sele��o, uma pergunta cadastrada para exclus�o.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Descri��o da Pergunta
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Descri��o da Pergunta</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe a descri��o da pergunta cadastrada.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar perguntas X �tem de checklist na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o o assunto.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Pergunta 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Pergunta</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode selecionar na caixa de sele��o uma pergunta cadastrada.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, consultar perguntas X �tem de checklist na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode consultar selecionando na caixa de sele��o o assunto.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Pergunta 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Pergunta</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode consultar selecionando na caixa de sele��o uma pergunta cadastrada.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Selecione uma empresa na caixa de sele��o, caso o usu�rio queira ativar uma pessoa f�sica.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Pergunta 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Pergunta</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Selecione uma empresa na caixa de sele��o, caso o usu�rio queira ativar uma pessoa f�sica.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'Tipo de Template
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Tipo de Template</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o o tipo de template. O Tipo de template pode ser: SA , SA Complementar, Papel de Trabalho e  Relat�rio de Auditoria.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'T�pico do Documento 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>T�pico do Documento</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode selecionar na caixa de sele��o um t�pico do documento. O t�pico do documento pode ser por exemplo, a base legal dentro da SA. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Campos Chave
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Campos Chave</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o o campo chave. Caso queira adicionar no texto no campo descri��o a palavra chave selecionado, basta clicar no Link adicionar ao texto. O campo chave serve para que o sistema pegue automaticamente da base de dados, informa��es desejadas sem que o usu�rio precise digit�-las.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Templates Cadastrados
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Templates Cadastrados</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Serve apenas para consultar os templates j� cadastrados na base de dados.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Descri��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Descri��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar o conte�do do template.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'Tipo de Template
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Tipo de Template</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o o tipo de template.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'T�pico do Documento 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>T�pico do Documento</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode selecionar na caixa de sele��o um t�pico do documento.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Campos Chave
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Campos Chave</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o o campo chave. Caso queira adicionar no texto no campo descri��o a palavra chave selecionado, basta clicar no Link adicionar ao texto.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Templates Cadastrados
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Templates Cadastrados/a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o o template cadastrados.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Descri��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Descri��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar alguma descri��o se desejar.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'Tipo de Template
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Tipo de Template</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o o tipo de template.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'T�pico do Documento 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>T�pico do Documento</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode selecionar na caixa de sele��o um t�pico do documento.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Templates Cadastrados
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Templates Cadastrados</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o o template que deseja alterar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Descri��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Descri��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar alguma descri��o se desejar.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na imagem ao lado do campo o n�mero do processo que queira desbloquear.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, exportar processo da base de dados para o m�dulo notebook.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado para exportar clicando na imagem ao lado do campo. </td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Importar Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Importar Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve clicar no Bot�o Procurar para escolher o arquivo (auditoria.mdb que foi unificado ap�s a auditoria)  que queira importar. Ap�s isto, o sistema exibe essa tela, para escolher o arquivo para importar.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'Data de Inicio
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Data de Inicio</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar a data inicio do processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Data Fim
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Data Fim</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar a data fim do processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
	
		'Unidade Auditada
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Unidade Auditada</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar o bot�o de sele��o RNML ou UP, para o sistema exibir na caixa de sele��o as unidades existentes. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Tipo de Auditoria
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Tipo de Auditoria</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar o bot�o de sele��o ordin�ria ou extraordin�ria, para marcar o tipo de Auditoria.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Auditor Respons�vel
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Auditor Respons�vel</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar o auditor respons�vel.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Fun��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#6 name=#6>Fun��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar a fun��o do auditor.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela para alterar ou bloquear o processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data de Inicio
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Data de Inicio</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar a data inicio do processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Data Fim
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data Fim</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar a data fim do processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
	
		'Unidade Auditada
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Unidade Auditada</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode selecionar o bot�o de sele��o RNML ou UP, para o sistema exibir na caixa de sele��o as unidades auditadas para sele��o referente ao bot�o selecionado. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Tipo de Auditoria
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Tipo de Auditoria</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode selecionar o bot�o de sele��o ordin�ria ou extraordin�ria, para marcar o tipo de Auditoria.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Auditor Respons�vel
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#6 name=#6>Auditor Respons�vel</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar o auditor respons�vel.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Fun��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#7 name=#7>Fun��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar a fun��o do auditor.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
			
		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela para excluir o processo.</td>" & vbcrlf 
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

		'Auditor Respons�vel
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#6 name=#6>Auditor Respons�vel</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio exibe o auditor respons�vel.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Fun��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#7 name=#7>Fun��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio exibe a fun��o do auditor.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Selecione uma empresa na caixa de sele��o, caso o usu�rio queira ativar uma pessoa f�sica.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data de Inicio
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Data de Inicio</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Selecione uma empresa na caixa de sele��o, caso o usu�rio queira ativar uma pessoa f�sica.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Data Fim
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data Fim</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Selecione uma empresa na caixa de sele��o, caso o usu�rio queira ativar uma pessoa f�sica.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
	
		'Unidade Auditada
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Unidade Auditada</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Selecione uma empresa na caixa de sele��o, caso o usu�rio queira ativar uma pessoa f�sica.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Tipo de Auditoria
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Tipo de Auditoria</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Selecione uma empresa na caixa de sele��o, caso o usu�rio queira ativar uma pessoa f�sica.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Auditor Respons�vel
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#6 name=#6>Auditor Respons�vel</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Auditor respons�vel cadastrada  para o processo selecionado.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Fun��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#7 name=#7>Fun��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Fun��o cadastrada  para o processo selecionado.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'Unidade 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Unidade</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode selecionar na caixa de sele��o a unidade desejada para pesquisa.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Auditor
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Auditor</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar o nome do auditor para pesquisar ou clicar na imagem ao lado do campo para pesquisar o auditor.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar as informa��es do auditor</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'Conselho Regional 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Conselho Regional</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar o conselho regional do auditor.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'N�mero do Registro 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>N�mero do Registro</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar o n�mero do registro do auditor.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'UF 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>UF</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar na caixa de sele��o a unidade federativa do auditor.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
	
		'Fun��o 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Fun��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar a fun��o do auditor.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela para excluir o processo.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
			
		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela para excluir o processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'N�mero da SA
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>N�mero da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Selecione uma empresa na caixa de sele��o, caso o usu�rio queira ativar uma pessoa f�sica.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode cadastrar a data da SA.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Tipo da SA
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Tipo da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode selecionar o bot�o de sele��o normal ou complementar, para marcar o tipo de SA.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela para alterar o processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'N�mero da SA
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>N�mero da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar a SA que deseja alterar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar a data da SA.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Tipo da SA
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Tipo da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode selecionar o bot�o de sele��o normal ou complementar, para marcar o tipo de SA.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela para excluir o processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'N�mero da SA
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>N�mero da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode selecionar qual SA ele deseja excluir.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
			
		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela para excluir o processo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'N�mero da SA
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>N�mero da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode selecionar qual SA ele deseja consultar.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
			
		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'N�mero da SA 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>N�mero da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Ap�s selecionar o n�mero do processo, o usu�rio deve selecionar na caixa de sele��o o n�mero da SA.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Ap�s selecionar a SA, o usu�rio deve selecionar na caixa de sele��o o assunto.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'�tem
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>�tem</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Ap�s selecionar os campos anteriores, o sistema exibe uma listagem dos Itens, referente aos campos selecionados.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
			
		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'N�mero da SA 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>N�mero da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Ap�s selecionar o n�mero do processo, o usu�rio deve selecionar na caixa de sele��o o n�mero da SA.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Ap�s selecionar a SA, o usu�rio deve selecionar na caixa de sele��o o assunto.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'�tem
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>�tem</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Ap�s selecionar os campos anteriores, o sistema exibe uma listagem dos Itens, referente aos campos selecionados.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, exibir todos os �tens da SA dispon�veis para o papel de trabalho.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
				
		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'N�mero da SA 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>N�mero da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Ap�s selecionar o n�mero do processo, o usu�rio deve selecionar na caixa de sele��o o n�mero da SA.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'�rea
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>�rea</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Ap�s selecionar o n�mero da SA, o usu�rio pode selecionar na caixa de sele��o a �rea.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Assunto
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Assunto</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Ap�s selecionar o n�mero da SA, o usu�rio pode selecionar na caixa de sele��o o assunto.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'�tens
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>�tens</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Ap�s selecionar os campos, o sistema exibe uma listagem dos Itens da SA, referente aos campos selecionados.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar o coment�rio na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o n�mero do processo, referente ao item selecionado na tela anterior.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'�tem
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>�tem</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o �tem, referente ao item selecionado na tela anterior.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Ir para Relat�rio 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Ir para Relat�rio </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode selecionar op��o se deseja que esse coment�rio v� para o relat�rio.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Ir para Ata de Auditoria 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Ir para Ata de Auditoria </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode selecionar na nesta caixa de texto se deseja que esse coment�rio v� para a Ata de Auditoria.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Coment�rio 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Coment�rio </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Para inserir um coment�rio, o usu�rio deve digitar neste campo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Recomenda��o  
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#6 name=#6>Recomenda��o  </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Para inserir uma recomenda��o, o usu�rio deve digitar neste campo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Anexos  
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#7 name=#7>Anexos </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Para informar se existem anexos existentes a esse coment�rio.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar um coment�rio.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o n�mero do processo, referente ao item selecionado na tela anterior.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'�tem
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>�tem</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o �tem, referente ao item selecionado na tela anterior.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Ir para Relat�rio 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Ir para Relat�rio </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode selecionar op��o se deseja que esse coment�rio v� para o relat�rio.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Ir para Ata de Auditoria 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#4 name=#4>Ir para Ata de Auditoria </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode selecionar na nesta caixa de texto se deseja que esse coment�rio v� para a Ata de Auditoria.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
			
		'Coment�rio 
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#5 name=#5>Coment�rio </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Para inserir um coment�rio, o usu�rio deve digitar neste campo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Recomenda��o  
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#6 name=#6>Recomenda��o  </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Para inserir uma recomenda��o, o usu�rio deve digitar neste campo.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Anexos  
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#7 name=#7>Anexos </a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Para informar se existem anexos existentes a esse coment�rio.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, consultar o coment�rio na base de dados. O sistema exibe os coment�rios referentes ao �tem selecionado na tela anterior</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o n�mero do processo, referente ao item selecionado na tela anterior.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'N�mero da SA
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>N�mero da SA</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O sistema exibe o n�mero da SA, referente ao item selecionado na tela anterior.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, cadastrar o �tem da ata de Auditoria na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar o �tem da ata de Auditoria na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
				
		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, excluir o �tem da ata de Auditoria na base de dados.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
				
		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os �tens da ata de auditoria na p�gina. </td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, consultar os �tens da ata ou imprim�-la.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
				
		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, gerar um relat�rio.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela para preenchimento. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'N�mero do Of�cio
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>N�mero do Of�cio</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar o n�mero do of�cio.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Data do Of�cio
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data do Of�cio</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar a data do of�cio.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, alterar um relat�rio.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
		
		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela para alterar. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'N�mero do Of�cio
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>N�mero do Of�cio</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar o n�mero do of�cio.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Data do Of�cio
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data do Of�cio</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar a data do of�cio.</td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, excluir um relat�rio.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
			
		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibir� as informa��es contidas no processo de auditoria  selecionado. </td>" & vbcrlf 
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
		'Cabe�alho
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
	    strRetorno = strRetorno & "<td  align='center' class='texto-titulo-HelpOnline'><a class='texto-titulo-HelpOnline' href=#0 name=#0></a>Objetivo da Tela</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr class='lin-0-1'>"
		strRetorno = strRetorno & "<td>Esta tela tem como objetivo, consultar ou imprimir um relat�rio de auditoria.</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "</table>"
		
		'Inicio da Tabela
		strRetorno = strRetorno & "<table width='90%' border='0' align='center' cellspacing='2' cellpadding='2'>"
		strRetorno = strRetorno & "<tr>"
    	strRetorno = strRetorno & "<td align='center' colspan=2 class='texto-titulo-HelpOnline'>Help Online</td>"
		strRetorno = strRetorno & "</tr>"
		strRetorno = strRetorno & "<tr>"
		strRetorno = strRetorno & "<td width='20%' valign=top class='SubtituloHelpOnline'>Campo</td>"
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela para preenchimento. </td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'N�mero do Of�cio
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>N�mero do Of�cio</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar o n�mero do of�cio para consultar.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 
		
		'Data do Of�cio
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#3 name=#3>Data do Of�cio</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar a data do of�cio para consultar.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'Arquivo para Importa��o
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>Arquivo para Importa��o</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>Selecione uma empresa na caixa de sele��o, caso o usu�rio queira ativar uma pessoa f�sica.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data do Parecer
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Data do Parecer</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio deve digitar a data do parecer.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
			
		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
		strRetorno = strRetorno & "</tr>" & vbcrlf 

		'Data do Parecer
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#2 name=#2>Data do Parecer</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode alterar a data do parecer.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"
			
		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
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
		'Cabe�alho
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
		strRetorno = strRetorno & "<td width='80%' class='SubtituloHelpOnline'>Descri��o</td>"
		strRetorno = strRetorno & "</tr>"

		'N�mero do Processo
		strRetorno = strRetorno & "<tr>" & vbcrlf 
		strRetorno = strRetorno & "<td width='20%' valign=top class='NomeCampoHelpOnline'><a class='link_help' href=#1 name=#1>N�mero do Processo</a></td>" & vbcrlf 
		strRetorno = strRetorno & "<td width='80%' class='lin-1-1'>O usu�rio pode digitar ou selecionar o processo desejado clicando na imagem ao lado do campo. Ap�s selecionado o n�mero do processo, o sistema exibe os dados na tela.</td>" & vbcrlf 
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
		strRetorno = "N�o h� suporte para este campo."
	end select
	help = strRetorno
	
end function

%>