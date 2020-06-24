<%
if Application("Indicador_notebook") = "S" then
	'NOTEBOOK
	if Len(trim(Session("Indicador_notebook"))) = 0 then
		Response.write "Sessão expirou. Se logue novamente no sistema"
		Response.End
	end if
else
	'REDE
	if Len(trim(session("usuario_rede"))) = 0 or Len(trim(session("menu"))) = 0 then
		Response.Redirect("Index.asp?NumeroErro=1")
	end if
end if
%>