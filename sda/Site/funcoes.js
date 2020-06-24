// onkeypress="DataSemBarra();"
function DataSemBarra()//Só aceita numero e enter
{
	var tecla = event.keyCode;
	if (tecla > 47 && tecla < 58) // numeros de 0 a 9
		return true;
	else {
	if (tecla != 13) // Enter
		event.keyCode = 0;
    else
		return true;
	}
}


//OnKeyUp="mascara_data(this)"
function mascara_data(mydata){ 
              
    if (event.keyCode!=8){ //BackSpace
  	//trata barra do dia
  	if (mydata.value.length == 2){ 
  	    mydata.value = mydata.value + '/'; 
  	} 
	if (mydata.value.length == 3 && mydata.value.substring(2,3) != '/'){
	    mydata.value = mydata.value.substring(0,2) + '/' + mydata.value.substring(2,3); 
	} 

  	//trata barra do mês
  	if (mydata.value.length == 5){ 
  	    mydata.value = mydata.value + '/'; 
  	} 
  	if (mydata.value.length == 6 && mydata.value.substring(5,6) != '/'){
  	    mydata.value = mydata.value.substring(0,5) + '/' + mydata.value.substring(5,6); 
  	} 

  	if (mydata.value.length == 10){ 
  	    verifica_data(mydata); 
  	} 
    }
} 

function ValidaTamanho(objTexto,tamanhoMax) {
	if (objTexto.value.length>tamanhoMax) {
		alert("Este campo está excedendo o tamanho máximo de "+(tamanhoMax)+" caracteres!");
		objTexto.value=objTexto.value.substring(0,tamanhoMax);
	}
}

function AbreJanela(pagina){
	if(pagina=='ProcessoBuscar.asp')
	{
		if(document.frm.NumProcesso.value==''){
			msg = window.open(pagina,"Janela2","width=790,height=570,scrollbars=yes,top=0,left=0");
		}
		else{
			document.frm.submit();
		}
	}
	else
		msg = window.open(pagina,"Janela2","width=790,height=570,scrollbars=yes,top=0,left=0");
}

function verifica_resolucao(){
	if (screen.height > 600){
		document.all.Td_Desenvolvimento.height='210px';
	}else{
		document.all.Td_Desenvolvimento.height='260px';
	}
}

function PassarProximoFoco(NameControlFocus){
	var objNextControl;
	var tecla = event.keyCode;
	
	objNextControl = eval("document.getElementsByName('" + NameControlFocus +  "');");
	objNextControl(0).focus();
}


