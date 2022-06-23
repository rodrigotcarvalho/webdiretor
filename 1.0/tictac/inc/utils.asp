<%
function pontinhos(palavra,numero,quebraPalavra)
'Reduzir o tamanho de um texto
	if len(palavra) > numero then
	
		texto = left(palavra,numero)

		if quebraPalavra = "S" then
			pontinhos = left(texto,numero) & "..."		
		else		
			procura = instrrev(texto,chr(32))
				
			pontinhos = left(texto,cint(procura) - 1) & "..."
		end if	
	
	else
	
		pontinhos = palavra
	
	end if

end function
%>