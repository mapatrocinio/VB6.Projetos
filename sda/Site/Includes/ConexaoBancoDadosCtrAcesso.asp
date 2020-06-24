<%
'on error resume next
dim CnCrtAcesso

set ObjCon = Server.CreateObject("bd_sda.ClsBD")
if err.number <> 0 then TrataErro

'Quando chamamos o RetornaStrCon, podemos passar 2 parâmentros sendo que o primeiro 
'é o tipo de conexão ou seja:
'  - 1 = >  conexão com o BD auditoria SQL SERVER
'  - 2 = >  conexão com o BD Access
'  - 3 = >  conexão com o BD CtrAcesso SQL SERVER

'E o segundo é o path para o banco em access na importação, exportação ou unificação
'dos papéis de trabalho

strconexao = ObjCon.RetornaStrCon(3,"")
if err.number <> 0 then TrataErro



set CnCrtAcesso = Server.CreateObject("ADODB.CONNECTION")
if err.number <> 0 then TrataErro
'CnCrtAcesso.Open "Provider=SQLOLEDB.1;Password=usracesso;Persist Security Info=True;User ID=acessousr;Initial Catalog=ctracesso;Data Source=xinm01s"
CnCrtAcesso.Open strconexao
set ObjCon = nothing


%>