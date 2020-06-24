<%
'**********************************************************************
'ROTINA PARA CRIACAO DE DIRETORIO
'PARAMETROS:
'Dir -> NOME DO DIRETORIO
'RETORNO:
'NENHUM

'**********************************************************************

Sub FileMakeDir(Dir)
	on error resume next
	Dim ArquivoObjeto
	Dim NomeCompleto
		
   	NomeCompleto = Dir
	
	Set ArquivoObjeto = Server.CreateObject("Scripting.FileSystemObject")
	
	
	If Not ArquivoObjeto.FolderExists(NomeCompleto) Then
		Call ArquivoObjeto.CreateFolder(NomeCompleto)
	End If

	Set ArquivoObjeto = Nothing

End Sub
%>