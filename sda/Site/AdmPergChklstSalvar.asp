<!-- #include file="Includes/ValidaSessao.asp" -->
<!-- #include file="Includes/Funcoes.asp" -->
<!--#include file="Includes/ConexaobancoDados.asp"-->
<!--#include file="Includes/Trataerro.asp"-->
<%
'----------------------------
'Funcao para retornar a string de permissões
'dentro de uma função
'----------------------------
Dim strAcoesPC
Dim strAcao


	strAcao = "INC"
	strAcoesPC = retornar_acoes("SDA", _
								 "MOD_AUD", _
								 "FC_ADM_PC_CAD")

valida_acesso_acao strAcoesPC, _
				   strAcao
'FIM ----------------------------


dim Acao 
dim Aux
dim sql
dim lStrErro
dim lintErro
dim lstrMensagem
dim FlagSalvar
dim i

dim Rs
dim RsAssunto
dim RsTemplate
dim RsCheckList

dim TextoAcao

dim VetChklst
dim seq_assunto_aux
dim seq_assunto
dim seq_item_sa
dim seq_checklist
'dim seq_topico_doc
'dim seq_tipo_documento
on error resume next

lintErro = 0

Acao = Request("Acao")
if Request("Acao") = "I" then
	TextoAcao = "Cadastrar"
	NomeTela = "perguntar_checklist_incluir"			
elseif Request("Acao") = "A" then
	TextoAcao = "Alterar"
	NomeTela = "perguntar_checklist_alterar"			
end if

Aux = Request("Aux")
seq_assunto_aux = Request("seq_assunto_aux")
seq_assunto = Request("SelAssunto")
seq_item_sa = Request("SelItemSa")
seq_checklist = Request("SelChecklist")

FlagSalvar = Request("FlagSalvar")

seq_tipo_documento = 1
seq_topico_doc = 3

'Esse flag seta que se houver tanto assunto, quando template e checklist, que poderá executar o cadastro.
if FlagSalvar = "" then
	FlagSalvar = 0
end if

if Aux <> "" then
	if FlagSalvar = 0 then
		lintErro = -99
		lStrErro = "É necessário ter uma pergunta e um ítem de checklist associado a um assunto."
	end if
end if

if Aux <> ""  and lintErro = 0 then
	sqlAux = ""
	'Como o checklist tem o mesmo nome de variável, jogar para um cursor todos os checks que foram selecionados
	VetChklst = split(seq_checklist,",")
	for i = 0 to ubound(VetChklst)
		sql = "spAdmPergChklstSalvar " & seq_assunto & "," & VetChklst(i) & "," & seq_item_sa &  ",'" & Session("Usuario_rede") & "','" & Acao & "'"
		cn.execute sql
		if err.number <> 0 then 
			'vai salvando, se der erro de PK, ignorar.
			if err.number <> -2147217873 then
				TrataErro
			end if
		end if
	next
	for i = 0 to ubound(VetChklst)
		sqlAux  = sqlAux &  VetChklst(i) & ","
	next
	if sqlAux <> "" then
		sqlAux = mid(sqlAux,1,len(sqlAux)-1)
	end if 
	err.Clear
	'remove os relacionamentos que não foram selecionados do banco de dados
	sql = "spAdmPergChklstRemover " & seq_assunto & "," & seq_item_sa & ",'" & sqlAux & "'"
	cn.execute sql
	if err.number <> 0 then 
		
		if err.number = -2147217873 then
			lintErro = -99 
			lStrErro = lStrErro & "Não é possível desassociar essa Pergunta de algum(ns) ítem(ns) de Checklist selecionados.<BR>Já existem comentários associados a esse ítem de checklist.<br>"
		else	
			TrataErro 
		end if
	end if
	if lintErro = 0 then
		Response.Redirect "ResultAdmPergChklst.asp?Acao="&Acao&"&SelAssunto="&seq_assunto&"&SelItemSa="&seq_item_sa
	end if
end if

set RsAssunto = server.CreateObject("ADODB.RECORDSET")
sql = "spListaAssunto"
RsAssunto.Open sql,cn,0,1



%>
<html>
<head>
<title>Sistema de Auditoria</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<Script LANGUAGE="JavaScript" src="funcoes.js"></Script>
<script language="JavaScript" src="funcoesHelp.js"></script>
<link rel="stylesheet" href="padraocss2.css" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
	var HideArray=['HideDiv1','HideDiv2'];
//-->
</SCRIPT>
</head>

<body bgcolor="#ccc1a1" text="#000000" background="imagens/azul_fundo1.gif"><!-- Início do Menu -->
<form name="frm" method="post" action="AdmPergChklstSalvar.asp">
<!-- Início do Menu -->
<!-- #include file="Includes/Cabecalho.asp" -->
<!-- Fim do Menu -->


<!-- Início do Tab -->
<!-- #include file="includes/menu_content.asp" -->
<!-- Fim da Tab -->

<!-- Início do Tab -->
<!--
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="44">
  <tr> 
    <td width="89"><IMG height=45 src="imagens/tab_esq.gif" width=89></td>
    <td background="imagens/tab_main.gif" class="texto" align="middle">&nbsp;</td>
    <td width="89"><IMG height=45 src="imagens/tab_dir.gif" width=89></td>
  </tr>
</table>
-->
<!-- Fim da Tab --><!-- Início do Tela de Desenvolvimento -->
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="51" height="35"><IMG height=35 src="imagens/main_esq_top.gif" width=51></td>
    <td width="100%" background="imagens/main_center_top.gif" height="17">&nbsp;</td>
    <td width="51" height="35"><IMG height=35 src="imagens/main_dir_top.gif" width=51></td>
  </tr>
  <tr valign="top"> 
    <td background="imagens/main_esq_center.gif" width="51">&nbsp;</td>
    <td id="Td_Desenvolvimento" width="100%" bgcolor="#ffffff" class="texto">
    <!--- ***************************Tabela de trabalho ************************************--->
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<TR height="22">
          <TD background="imagens/azul_claro.gif" align="center" class=azul colSpan=3><STRONG><%=TextoAcao%> Perguntas x Ítens de CheckList</STRONG></TD>
          </TR>
        <tr><td class="texto" colspan=3>&nbsp;</td></tr>
        <tr>
  	        <td colspan=3 align="right"><a href="javascript:help_online('<%=NomeTela%>',0);"><img src="imagens/interrogacao.gif" alt="Descrição do objetivo da tela" width="16" height="16" border=0></a></td>
		</tr>                		                
        <%if lintErro <> 0 then %>
		<tr><td class="texto" colspan="3" align="center"><font color="red"><%=lStrErro%></font></td></tr>
        <tr><td class="texto" colspan=3>&nbsp;</td></tr>
		<%
	end if%>
		<tr height="22"> 
			<td class="texto" width="10%"><a class="link_Campo" href="javascript:help_online('<%=NomeTela%>',1);">Assunto:</td>			
			<td colspan=2>
				<span id="HideDiv1" style="position:relative;">
				<SELECT class=texto name="SelAssunto" style="HEIGHT: 22px; WIDTH: 50%" onchange="javascript:document.frm.submit();">
				<option value="0">&lt;-- Assunto --&gt;</option>
				<%
				if not RsAssunto.EOF then
					FlagSalvar = 1
					'if seq_assunto = "" then
					'	seq_assunto = RsAssunto("seq_assunto")
					'end if
				end if
				do while not RsAssunto.EOF %>
					<OPTION <%if cint(seq_assunto) = cint(RsAssunto("seq_assunto")) then%> selected <%end if%> value="<%=RsAssunto("seq_assunto")%>"><%=RsAssunto("descr_assunto")%></option>  
				<%
				RsAssunto.MoveNext
				loop
				RsAssunto.Close
				%>
				</SELECT>
				</span>
			</td>
        </tr>
        <!-- Se não existe template cadastrado para esse assunto, mensagem de erro -->
    <%
    set RsTemplate = server.CreateObject("ADODB.RECORDSET")		
		if seq_assunto <> "" and seq_assunto <> 0 then
			sql = "spListaPergunta " & seq_assunto
			RsTemplate.Open sql,cn,0,1
			if not RsTemplate.eof then 
				if seq_item_sa = "" then
					seq_item_sa = RsTemplate("seq_item_sa")
				end if%>
			  <tr height="22"> 
				<td class="texto" width="10%"><a class="link_Campo" href="javascript:help_online('<%=NomeTela%>',2);">Pergunta:</td>			
				<td colspan=2>
					<span id="HideDiv2" style="position:relative;">
					<SELECT class=texto name="SelItemSa" style="HEIGHT: 22px; WIDTH: 100%" onchange="javascript:document.frm.submit();">
						<%
						do while not RsTemplate.EOF %>
						<OPTION <%if cint(seq_item_sa) = cint(RsTemplate("seq_item_sa")) then%> selected <%end if%> value="<%=RsTemplate("seq_item_sa")%>"><%=RsTemplate("descr_item_sa")%></OPTION> 
						<%
						RsTemplate.MoveNext
						loop 
						RsTemplate.Close%>
					</SELECT>
					</span>
				</td>
			  </tr>
				<%
			else
				FlagSalvar = 0
				%>
						<span id="HideDiv2" style="position:relative;">
						</span>
				
				<%
			end if
			%>
	      <tr><td colspan=3>&nbsp;</td></tr>
        <TR>
			<TD colspan=3 width="97%" class="texto" align="center">
				<TABLE border=0 cellPadding=1 cellSpacing=1 width="100%">
				<TR>
					<TD>&nbsp;</TD>
					<TD rowspan="3" nowrap class="texto" align="center" style="WIDTH: 20%" width="20%"><b>Ítens de CheckList</b></TD>
					<TD style="WIDTH: 40%" width="40%">&nbsp;</TD>
				</TR>
				<TR>
					<TD background="imagens/azul_fundo1.gif" 
						  style="HEIGHT: 1px" height=1></TD>
					<TD height=5 style="HEIGHT: 1px" background="imagens/azul_fundo1.gif">
					</TD>
				</TR>
				<TR>
					<TD>&nbsp;</TD>
					<TD>&nbsp;</TD>
				</TR>
				</TABLE>  
			</TD>
			</TR>
        <%
			if FlagSalvar = 1 then
				set RsCheckList = server.CreateObject("ADODB.RECORDSET")  
				sql = "spListaChecklist " & seq_assunto
				RsCheckList.Open sql, cn,0,1

				if not RsCheckList.EOF then
						set Rs = server.CreateObject("ADODB.RECORDSET")
						sql = "spListaPerguntaChecklist " & seq_assunto & "," & seq_item_sa
						rs.Open sql, cn
						
					Do while not RsCheckList.EOF
						'Esse filtro serve para verificar se o ítem de checklist está associado a Pergunta
						'para esse determinado assunto.
						'if not rs.EOF then
							rs.filter  = "seq_checklist =" & RsCheckList("seq_checklist")				
						'end if			
			      %>
			      <tr> 
					<td class="texto" width="97%" 
			          colSpan=2>				 
						<li><%=RsCheckList("descr_checklist")%></li>     
					</td>
					<td width="3%" align="center"><INPUT <%if not rs.eof then%> checked <%end if%> name="SelCheckList" type=checkbox value="<%=RsCheckList("seq_checklist")%>">
					</td>
			      </tr>
			      <%
						rs.filter  = ""
					RsCheckList.MoveNext
					loop
					if not rs.EOF then
						rs.Close
					end if
				else
					FlagSalvar = 0
				end if 
			end if
	else%>
		<span id="HideDiv2" style="position:relative;">
		</span>
	
	<%
	end if%>
        </table>
    </td>
    <td background="imagens/main_dir_center.gif" width="51">&nbsp;</td>
  </tr>
  <tr> 
    <td height="28" width="51"><IMG height=28 src="imagens/main_esq_low.gif" width=51></td>
    <td width="100%" background="imagens/main_center_low.gif" height="28">&nbsp;</td>
    <td width="51" height="28"><IMG height=28 src="imagens/main_dir_low.gif" width=51></td>
  </tr>
</table>
<!-- Fim da Tela de Desenvolvimento --><!-- Início do Roda pé -->
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="62" height="33"><IMG src="imagens/low_esq.gif"></td>
    <td background="imagens/low_center.gif" valign="center" >
      <div align="center"><!-- Início da Imagem do Botão de Navegação-->
        <table cellspacing="0" cellpadding="0" bordercolor="#009999" width="133" height="22">
          <tr> 
            <td rowspan="3" height="7" width="6" align="right"> <IMG height=20 src="imagens/cornerA.gif" width=6></td>
            <td height="3" background="imagens/fio1a.gif" width="111"></td>
            <td rowspan="3" height="7" width="14" align="left"><IMG height=20 src="imagens/cornerB.gif" width=6></td>
          </tr>
          <tr> 
            <td bgcolor="#f6c6b5" height="13" width="111" background="imagens/amarelo.gif"><!-- Início do Botão de Navegação-->
              <table class="botaoverde" border="0" cellspacing="0" cellpadding="0" align="center" width="107" height="5">
                <tr> 
                  <td width="100%" nowrap height="5" bgcolor="#f6c6b5" background="imagens/amarelo.gif"> 
                    <font size="1"><A class="link_botao_rodape" href="auditoria.asp">Página Principal</A></font>
                    &nbsp;::&nbsp;<font size="1"><A class="link_botao_rodape" href="javascript:SubmeterForm();">Salvar</A></font></td>
                </tr>
              </table>
			  <!-- Fim do Botão de Navegação-->
            </td>
          </tr>
          <tr> 
            <td background="imagens/fio2a.gif" width="111" height="4"></td>
          </tr>
        </table>
        <!-- Fim da Imagem do Botão de Navegação-->
      </div>
    </td>
    <td width="62" height="33"><IMG src="imagens/low_dir.gif"></td>
  </tr>
</table>
<!-- Fim do Roda pé -->
<input type="hidden" name="Aux">
<input type="hidden" name="Acao" value="<%=Acao%>">
<input type="hidden" name="FlagSalvar" value="<%=FlagSalvar%>">
<input type="hidden" name="seq_assunto_aux" value="<%=seq_assunto%>">
</form>
</body>
</html>
<SCRIPT LANGUAGE=javascript>
<!--
function SubmeterForm(){
	document.frm.Aux.value='Aux';
	document.frm.submit();
}
verifica_resolucao();
//-->
</SCRIPT>

<%
set Rs = nothing
set RsAssunto = nothing
set RsCheckList = nothing
set RsTemplate = nothing
set cn = nothing
%>
