function Salvar()
{
	Frm=document.F;
	if(Frm.Arquivo.value=='')
	{
		alert('Selecione o arquivo com a imagem a ser carregada.');
		return;
	}
	
	Frm.submit();
}
function MostraMensagem(Texto)
{
	if (Texto!='')
	{
		alert(Texto);
	}
}