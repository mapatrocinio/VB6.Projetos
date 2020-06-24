<%
'**********************************************************************
'ROTINA PARA VERIFICACAO SE O ARQUIVO EXISTE
'PARAMETROS:
'FileName -> NOME DO ARQUIVO
'RETORNO:
'TRUE SE EXISTIR

'**********************************************************************

Function FileExists(FileName)
'APAGA UM ARQUIVO DO SERVIDOR
'PARAMETROS:
'FileName -> NOME DO ARQUIVO
'RETORNO:
'TRUE SE EXISTIR

	Dim NomeCompleto
	Dim ArquivoObjeto
	
	If Left(Filename,2)="//" Or Left(Filename,2)="\\" Or Mid(Filename,2,1)=":" Then
	'CAMINHO ABSOLUTO
		NomeCompleto=FileName
	Else
	'CAMINHO RELATIVO AO SITE
		NomeCompleto = Server.MapPath("/" & FileName)
	End If
	Set ArquivoObjeto = Server.CreateObject("Scripting.FileSystemObject")
	FileExists=ArquivoObjeto.FileExists(NomeCompleto)
	Set ArquivoObjeto = Nothing
End Function
%>