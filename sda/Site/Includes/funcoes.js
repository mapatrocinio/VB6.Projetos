
function SubmeteForm(p_form, p_target){
	p_form.action = p_target
	p_form.submit()
}


function pressDate()//Só aceita numero, enter e barra
{
	var tecla = event.keyCode;
	if (tecla > 47 && tecla < 58) // numeros de 0 a 9
		return true;
	else {
	if (tecla != 13 && tecla != "47") // Enter e barra
		event.keyCode = 0;
    else
		return true;
	}
}


function pressNumber(campo) 
	{
	  tecla = window.event.keyCode;
	  if (tecla > 47 && tecla < 58) // numeros de 0 a 9
	  {
	         campo.value =campo.value  ;
	  }
   	  else
	  {
		if (tecla != 8 && tecla != 13) // backspace e Enter
			event.keyCode = 0;
		else
			return true;
	  }
	}


function ValidaSequencia(p_senha)
{
	var i = 1;
	var countAsc = 1;
	var countDec = 1;
	
	while (i < p_senha.length)
	{
		// caracteres em sequencia crescente
		if (eval(p_senha.charCodeAt(i-1)+1) == (p_senha.charCodeAt(i))) {
			if (countAsc >= 2)
			{
				return false;
			}	
			countAsc++;
		} else {
			countAsc = 1;
		}
		
		// caracteres em sequencia decrescente
		if (eval(p_senha.charCodeAt(i-1)-1) == (p_senha.charCodeAt(i))) {
			if (countDec >= 2)
			{
				return false;
			}	
			countDec++;
		} else {
			countDec = 1;
		}
		i++;	
	}
	return true;

}

function ValidaRepeticao(p_senha)
{
	
	var i = 1;
	var lastChar = p_senha.substring(0,1);
	var count = 1;
	
	while (i < p_senha.length)
	{
		// testa caracteres repetidos
		if (p_senha.substring(i,i+1) == lastChar) {
			count++
		}
		lastChar = p_senha.substring(i,i+1);
		i++;	
	}
	if (count >= 3)
	{
		return false;
	}
	return true;
}

function ValidaSenhaLogin(p_senha,p_login)
{
	
	if (p_senha == p_login)
	{
		return false;
	}
	
	var i = eval(p_login.length)
	var result = ""
	while (i >= 0)
	{
		//inverte login
		result = result + p_login.substring(i-1,i);
		i--;	
	}
	
	if (p_senha == result)
	{
		return false;
	}
	return true;
}

function valida_campo(p_campo, p_tipo, p_aceita_nulo, p_msg, opcional){
	if (p_tipo=='N'){ //NÚMERO
		if (p_campo.value.length==0){
			if (p_aceita_nulo==false){
				window.alert(p_msg)
				p_campo.focus()
				return false;
			}
		}else{ //DIGITOU O NÚMERO
			if (isNaN(p_campo.value)){
				window.alert(p_msg)
				p_campo.focus()
				return false;
			}
		}
	}
	if (p_tipo=='T'){ //TEXTO
		if (p_aceita_nulo==false){
			if (p_campo.value.length==0){
				window.alert(p_msg)
				p_campo.focus()
				return false;
			}
		}
	}
	
	if (p_tipo=='S'){ //SENHA
		if (p_campo.value.length<6)
		{
			window.alert("A senha dever ter no mínimo 6 caracteres")
			p_campo.focus()
			return false;
		}
		
		var numero = false;
		var letra = false;
		var i = 0;
		
		while (i < p_campo.value.length)
		{
			if (isNaN(p_campo.value.substring(i,i+1))) {
				letra = true
			} else {
				numero = true
			}
			i++;	
		}
		if((!numero) || (!letra))
		{
			window.alert("Sua senha deve obrigatoriamente ter letras e números.")
			p_campo.focus()
			return false;
		} 
		if(!ValidaRepeticao(p_campo.value))		
		{
			window.alert('Não é permitido letras ou números repetidos na senha. Ex.: "AAA" ou "222".');
			p_campo.focus();
			return false;
		}
		if(!ValidaSequencia(p_campo.value))		
		{
			window.alert('Não é permitido letras ou números em sequência na senha. Ex.: "ABC" ou "432".');
			p_campo.focus();
			return false;
		}
		if(!ValidaSenhaLogin(p_campo.value, opcional.value))		
		{
			window.alert("O campo senha não pode ser igual ao campo Login");
			p_campo.focus();
			return false;
		}
	}
	
	if (p_tipo=='D'){ //DATA
		if (p_campo.value.length==0){
			if (p_aceita_nulo==false){
				window.alert(p_msg)
				p_campo.focus()
				return false;
			}
		}else{ //DIGITOU A DATA
			if (!CriticaData(p_campo.value)){
				window.alert(p_msg)
				p_campo.focus()
				return false;
			}else{	// A DATA É VÁLIDA, VERIFICA SE FOI DIGITADA 
				//NO FORMATO DD/MM/YYYY
				if (p_campo.value.length!=10){
					window.alert(p_msg)
					p_campo.focus()
					return false;
				}
			}
		}
	}
	return true;
}

function CriticaData(strinput)	
{
	
	var barra1;
	var barra2;
	var parte1;
	var parte2;
	var partedia;
	var partemes;
	var parteano;
	var retorno;
	var restobisexto;
	var arrayultimodia;
	var datagerada;
	parte1=strinput.substr(0,3);
	barra1 = parte1.search("/");
	
	
	if (barra1 > 0 )
	{
			if (barra1 == 1) 
			{	
				partedia = "0" + parte1.substr(0,barra1);
			}
			else
			{
				partedia =  parte1.substr(0,barra1);
			}
			parte2=strinput.substr(barra1 + 1,3);
			barra2 = parte2.search("/");
			if (barra2 > 0)
			{
				
				if (barra2 == 1) 
				{	
					partemes = "0" + parte2.substr(0,barra2);
				}
				else
				{
					partemes =  parte2.substr(0,barra2);
				}
				parteano = strinput.substr(1 + barra1 + barra2 + 1);
				
				if (parteano.length == 2)
				{
					
					if (parseInt(parteano) > 50)
					{
						parteano = 19 + parteano;
					}
					else
					{
						parteano = 20 + parteano;
					}
					restobisexto = parteano % 4;
					if (restobisexto != 0)
					{
						arrayultimodia = new Array(31,28,31,30,31,30,31,31,30,31,30,31);
					}
					else
					{
						arrayultimodia = new Array(31,29,31,30,31,30,31,31,30,31,30,31);
					}
					if (partemes > 0) 
					{
						if (partemes <= 12)
						{
							if (partedia > 0)
							{
								if (partedia <= arrayultimodia[partemes-1])
								{
									datagerada = new Date(parseInt(parteano),parseInt(partemes), parseInt(partedia));
									if (datagerada)
									{
										retorno = true;
									}
								}
								else
								{
									retorno = false;
								}
							}
							else
							{
								retorno = false;
							}
						}
						else
						{
							retorno = false;
						}
					}
					else
					{
						retorno = false;
					}
							
				}
				else
				{
					if (parteano.length == 4)
					{	
						restobisexto = parteano % 4;
						if (restobisexto != 0)
						{
							arrayultimodia = new Array(31,28,31,30,31,30,31,31,30,31,30,31);
						}
						else
						{
							arrayultimodia = new Array(31,29,31,30,31,30,31,31,30,31,30,31);
						}
						if (partemes > 0) 
						{
							if (partemes <= 12)
							{
								if (partedia > 0)
								{
									if (partedia <= arrayultimodia[partemes-1])
									{
										datagerada = new Date(parseInt(parteano),parseInt(partemes), parseInt(partedia));
										if (datagerada)
										{
											retorno = true;
										}										
									}
									else
									{
										retorno = false;
									}
								}
								else
								{
									retorno = false;
								}
							}
							else
							{
								retorno = false;
							}
						}
						else
						{
							retorno = false;
						}
					}
					else 
					{
						// ano com tamanho diferente de 2 e 4
							retorno = false;
					}
				}
			}			
			else
			{
				// nao achou a segunda barra ou está na terceira posicao
				retorno = false;
			}
	}
	else
	{
		// nao achou a primeira barra ou está na primeira posicao
		retorno = false;
	}
	return retorno;
}
