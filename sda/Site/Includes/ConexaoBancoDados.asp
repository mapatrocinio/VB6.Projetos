
<%
'on error resume next
dim cn
dim MensagemErro
dim ObjCon
Dim strconexao 



set cn = Server.CreateObject("ADODB.CONNECTION")
if err.number <> 0 then TrataErro
set ObjCon = Server.CreateObject("bd_sda.ClsBD")
if err.number <> 0 then TrataErro


'Quando chamamos o RetornaStrCon, podemos passar 2 par�mentros sendo que o primeiro 
'� o tipo de conex�o ou seja:
'  - 1 = >  conex�o com o BD auditoria SQL SERVER
'  - 2 = >  conex�o com o BD auditoria SQL SERVER (M�dulo Notebook)
'  - 3 = >  conex�o com o BD CtrAcesso SQL SERVER
'E o segundo � o path para o banco em access na importa��o, exporta��o ou unifica��o
'dos pap�is de trabalho
If Application("Indicador_notebook") = "S" then
	strconexao = ObjCon.RetornaStrCon(2,"")
	
else
	strconexao = ObjCon.RetornaStrCon(1,"")
end if
if err.number <> 0 then TrataErro

'cn.Open "Provider=SQLOLEDB.1;Password=midesenv;Persist Security Info=True;User ID=mimontreal;Initial Catalog=auditoria;Data Source=rdes01s"

cn.Open strconexao

set ObjCon = nothing


%>