<%
'on error resume next
dim CnCrtAcesso

set ObjCon = Server.CreateObject("bd_sda.ClsBD")
if err.number <> 0 then TrataErro

'Quando chamamos o RetornaStrCon, podemos passar 2 par�mentros sendo que o primeiro 
'� o tipo de conex�o ou seja:
'  - 1 = >  conex�o com o BD auditoria SQL SERVER
'  - 2 = >  conex�o com o BD Access
'  - 3 = >  conex�o com o BD CtrAcesso SQL SERVER

'E o segundo � o path para o banco em access na importa��o, exporta��o ou unifica��o
'dos pap�is de trabalho

strconexao = ObjCon.RetornaStrCon(3,"")
if err.number <> 0 then TrataErro



set CnCrtAcesso = Server.CreateObject("ADODB.CONNECTION")
if err.number <> 0 then TrataErro
'CnCrtAcesso.Open "Provider=SQLOLEDB.1;Password=usracesso;Persist Security Info=True;User ID=acessousr;Initial Catalog=ctracesso;Data Source=xinm01s"
CnCrtAcesso.Open strconexao
set ObjCon = nothing


%>