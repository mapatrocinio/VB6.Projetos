<%

Response.Buffer = true

sub TrataErro()
	if MensagemErro = "" then
		MensagemErro = err.Description
	end if
	Response.Write "Erro: " & err.number & "<BR>"
	Response.Write "Descrição: " & MensagemErro & "<BR>"
	err.Clear
	err.number = 0
	Response.End

end sub


%>