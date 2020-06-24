<%
'Response.Write("nadaa")
if NumProcesso <> "" then
	if session("NumProcesso") <> NumProcesso then
		 session("NumProcesso") = NumProcesso
'		 Response.Write("seta sessao")
	end if
else
	if session("NumProcesso") <> "" and request.form("Hidden_Flag") = "" then
		NumProcesso = session("NumProcesso")
'		Response.Write("seta num processo")
	else
		NumProcesso = ""
		session("NumProcesso") = ""
	end if
end if
if seq_sa = "0-0" or seq_sa & "" = "" or len(seq_sa) < 3 then
	'combo da SA em branco, ou sa não é vállida
	'verifica se campo hidden da sa está preenchido da seleção
	if request("hidSaSel") <> "" then
		session("NumSA") = request("hidSaSel")
	end if
	if session("NumSA") <> "" and request.form("Hidden_Flag") = "" then
		seq_sa = session("NumSA")
	else
		if request("hidSaSel") <> "" then
			seq_sa = session("NumSA")
		else
			'session("NumSA") = seq_sa
		end if
	end if
else
	'Selecionou uma SA do combo, atualiza session sa
	if session("NumSA") <> seq_sa then
		 session("NumSA") = seq_sa
	end if
end if
%>


