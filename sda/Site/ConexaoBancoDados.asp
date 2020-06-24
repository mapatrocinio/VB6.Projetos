<%
dim cn
set cn = Server.CreateObject("ADODB.CONNECTION")
cn.Open "Provider=SQLOLEDB.1;Password=midesenv;Persist Security Info=True;User ID=mimontreal;Initial Catalog=auditoria;Data Source=sda"



%>