<%
'**********************************************************************
'CONVERTE UM BUFFER DE BYTES PARA STRING
'PARAMETROS:
'Bin -> BUFFER DE BYTES BINARIOS
'RETORNO:
'EQUIVALENTE STRING DO BUFFER
'DEPENDENCIAS:
'NENHUMA
'**********************************************************************
Function BinToStr(Bin)
    Dim Str
    Dim I
    
    Str =""
    For I = 1 to LenB(Bin)
	'CONVERTE CADA BYTE PARA CARACTERE
		Str = Str & Chr(AscB(MidB(Bin, I, 1)))
    Next
    BinToStr = Str
End Function
%>