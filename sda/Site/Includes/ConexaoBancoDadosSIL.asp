<%
'on error resume next
dim cn2

set cn2 = Server.CreateObject("ADODB.CONNECTION")
if err.number <> 0 then TrataErro
cn2.Open "Provider=SQLOLEDB.1;Password=midesenv;Persist Security Info=True;User ID=mimontreal;Initial Catalog=SIL;Data Source=rdes01s"

%>