<%
dim cnA


set cnA = Server.CreateObject("ADODB.CONNECTION")
if err.number <> 0 then TrataErro

set ObjCon = Server.CreateObject("bd_sda.ClsBD")
if err.number <> 0 then TrataErro

'Quando chamamos o RetornaStrCon, podemos passar 2 par�mentros sendo que o primeiro 
'� o tipo de conex�o ou seja:
'  - 1 = >  conex�o com o BD auditoria SQL SERVER
'  - 2 = >  conex�o com o BD Access
'  - 3 = >  conex�o com o BD CtrAcesso SQL SERVER
'E o segundo � o path para o banco em access na importa��o, exporta��o ou unifica��o
'dos pap�is de trabalho



strconexao = ObjCon.RetornaStrCon(2,Application("ParaAuditoria"))
if err.number <> 0 then TrataErro

'cnA.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.Mappath("bd\Exportacao") & "\auditoria.mdb"

cnA.Open strconexao
set ObjCon = nothing



%>