
<%
'on error resume next
dim cmd

set cmd = Server.CreateObject("ADODB.Command")
if err.number <> 0 then TrataErro

'cn.Open "Provider=SQLOLEDB.1;Password=midesenv;Persist Security Info=True;User ID=mimontreal;Initial Catalog=auditoria;Data Source=rdes01s"
cmd.ActiveConnection = cn
cmd.CommandType = 1 'adCmdText


%>