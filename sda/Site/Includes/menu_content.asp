<!--
<script src="configuracao.js"></script>
<script src="menuff.js"></script>
-->
<script language="JavaScript">

<!-- #include file="configuracao.js" --> 

</script>
<script language="JavaScript">

<!-- #include file="menuff.js" --> 

</script>

<SCRIPT LANGUAGE=javascript>

<%if Application("Indicador_notebook") = "N" then
		response.write(session("menu"))
	else%>



Menu1=new Array('Manutenção','','',3,24,80);

Menu1_1=new Array('Exportar Processo','ExportarProcesso.asp','',0,16,100);
Menu1_2=new Array('Importar Processo','ImportarProcesso.asp','',0,16,100);
Menu1_3=new Array('Unificação','Unificar.asp','',0,16,100);

Menu2=new Array('Processo','','',1,24,80);
Menu2_1=new Array('Consultar','ProcessoConsultar.asp','',0,16,80);

Menu3=new Array('SA','','',2,24,80);
Menu3_1=new Array('Cadastrar','SaSalvar.asp?Acao=I','',0,16,80);
Menu3_2=new Array('Consultar','SAConsultar.asp','',0,16,80);

Menu4=new Array('Checklist','','',2,24,80);
Menu4_1=new Array('Gerar','chklstSalvar.asp?Acao=I','',0,16,80);
Menu4_2=new Array('Consultar','ChklstConsultar.asp','',0,16,80);

Menu5=new Array('Papel de Trabalho','PapelTrabalhoListar.asp?Origem=0','',0,24,80);

Menu6=new Array('Ata','','',4,24,80);
Menu6_1=new Array('Cadastrar','AtaSalvar.asp?Acao=I','',0,16,80);
Menu6_2=new Array('Alterar','AtaSalvar.asp?Acao=A','',0,16,80);
Menu6_3=new Array('Excluir','AtaExcluir.asp','',0,16,80);
Menu6_4=new Array('Consultar','AtaConsultar.asp','',0,16,80);

Menu7=new Array('Relatórios','','',1,24,80);
Menu7_1=new Array('Auditoria','','',4,16,80);
Menu7_1_1=new Array('Gerar','RelatorioSalvar.asp?Acao=I','',0,16,150);
Menu7_1_2=new Array('Alterar','RelatorioSalvar.asp?Acao=A','',0,16,150);
Menu7_1_3=new Array('Excluir','RelatorioExcluir.asp','',0,16,150);
Menu7_1_4=new Array('Consultar','RelatorioConsultar.asp','',0,16,150);

Menu8=new Array('Parecer','','',4,24,80);
Menu8_1=new Array('Cadastrar','ParecerSalvar.asp?Acao=I','',0,16,80);
Menu8_2=new Array('Alterar','ParecerSalvar.asp?Acao=A','',0,16,80);
Menu8_3=new Array('Excluir','ParecerExcluir.asp','',0,16,80);
Menu8_4=new Array('Consultar','ParecerConsultar.asp','',0,16,80);

Menu9=new Array('Sair','javascript:close();','',0,24,80);

var NoOffFirstLineMenus=9; //Número de Itens do menu

<%end if%>

</SCRIPT>




