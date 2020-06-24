<%
dim cnA


set cnA = Server.CreateObject("ADODB.CONNECTION")
if err.number <> 0 then TrataErro

set ObjCon = Server.CreateObject("bd_sda.ClsBD")
if err.number <> 0 then TrataErro

'Quando chamamos o RetornaStrCon, podemos passar 2 parâmentros sendo que o primeiro 
'é o tipo de conexão ou seja:
'  - 1 = >  conexão com o BD auditoria SQL SERVER
'  - 2 = >  conexão com o BD Access
'  - 3 = >  conexão com o BD CtrAcesso SQL SERVER
'E o segundo é o path para o banco em access na importação, exportação ou unificação
'dos papéis de trabalho



strconexao = ObjCon.RetornaStrCon(2,Application("ParaAuditoria"))
if err.number <> 0 then TrataErro

'cnA.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.Mappath("bd\Exportacao") & "\auditoria.mdb"

cnA.Open strconexao
set ObjCon = nothing



%>