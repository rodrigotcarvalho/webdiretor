<%
FUNCTION mensagens_escolas(ambiente_escola,nivel,msg,tab,dados1,dados2,dados3)

SELECT CASE msg
'lan&ccedil;amento de notas de 1000 a 5999
case 1000	
	if ambiente_escola="boechat" or ambiente_escola="testeboechat"  then
		SELECT CASE dados2
			case "f$0"
				wrt = "O campo faltas do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"
				errou="faltas"
			case "f$1"
				wrt = "O campo faltas do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor que 255"				
				errou="faltas"
			case "p_av1$0"
				wrt = "O campo Peso da AV1 deve conter um n&uacute;mero inteiro"				
			case "p_av1$1"
				wrt = "O Peso da AV1 deve ser menor ou igual a 100"
			case "p_av2$0"
				wrt = "O campo Peso da AV2 deve conter um n&uacute;mero inteiro"				
			case "p_av2$1"
				wrt = "O Peso da AV2 deve ser menor ou igual a 100"				
			case "p_av3$0"
				wrt = "O campo Peso da AV3 deve conter um n&uacute;mero inteiro"				
			case "p_av3$1"
				wrt = "O Peso da AV3 deve ser menor ou igual a 100"
			case "p_av4$0"
				wrt = "O campo Peso da AV4 deve conter um n&uacute;mero inteiro"				
			case "p_av4$1"
				wrt = "O Peso da AV4 deve ser menor ou igual a 100"				
			case "p_av5$0"
				wrt = "O campo Peso da AV5 deve conter um n&uacute;mero inteiro"				
			case "p_av5$1"
				wrt = "O Peso da AV5 deve ser menor ou igual a 100"	
			case "smp"
				wrt = "A soma dos Pesos deve ser menor ou igual a 100"	
			case "av1$0"
				wrt = "O campo AV1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="av1"
			case "av1$1"
				wrt = "O campo AV1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual ao seu peso"
				errou="av1"
			case "av2$0"
				wrt = "O campo AV2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="av2"
			case "av2$1"
				wrt = "O campo AV2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual ao seu peso"				
				errou="av2"
			case "av3$0"
				wrt = "O campo AV3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="av3"
			case "av3$1"
				wrt = "O campo AV3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual ao seu peso"
				errou="av3"
			case "av4$0"
				wrt = "O campo AV4 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="av4"
			case "av4$1"
				wrt = "O campo AV4 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual ao seu peso"				
				errou="av4"
			case "av5$0"
				wrt = "O campo AV5 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="av5"
			case "av5$1"
				wrt = "O campo AV5 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual ao seu peso"		
				errou="av5"
			case "pr$0"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="pr"
			case "pr$1"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 100"	
				errou="pr"
			case "bon$0"
				wrt = "O campo Bon do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="bon"
			case "bon$1"
				wrt = "O campo Bon do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 100"
				errou="bon"
			case "rec$0"
				wrt = "O campo Rec do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="rec"
			case "rec$1"
				wrt = "O campo Rec do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 100"
				errou="rec"
		end select
		
'================================================================================================================================		
	elseif ambiente_escola="insa" or ambiente_escola="testeinsa"  then
		SELECT CASE dados2
			case "f"
				wrt = "O campo faltas do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"
				errou="faltas"		
			case "av1$0"
				wrt = "O campo Tr1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="tr1"
			case "av1"
				wrt = "O campo Tr1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="av1"
			case "tr2$a"
				wrt = "O campo Tr2 n&atilde;o deve ser utilizado"				
				errou="tr2"						
			case "tr2$0"
				wrt = "O campo Tr2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="tr2"
			case "av2"
				wrt = "O campo Tr2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"				
				errou="av2"
			case "tr3$a"
				wrt = "O campo Tr3 n&atilde;o deve ser utilizado"				
				errou="tr3"						
			case "tr3$0"
				wrt = "O campo Tr3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="tr3"
			case "av3"
				wrt = "O campo Tr3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="av3"
			case "te1$a"
				wrt = "O campo Te1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 7"				
				errou="te1"				
			case "te1$0"
				wrt = "O campo Te1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="te1"
			case "te1$1"
				wrt = "O campo Te1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"				
				errou="te1"
			case "te2$a"
				wrt = "O campo Te2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 7"				
				errou="te2"					
			case "te2$0"
				wrt = "O campo Te2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"			
				errou="te2"
			case "te2$1"
				wrt = "O campo Te2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="te2"
			case "te3$a"
				wrt = "O campo Te3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 7"				
				errou="te3"					
			case "te3$0"
				wrt = "O campo Te3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"		
				errou="te3"
			case "te3$1"
				wrt = "O campo Te3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="te3"
			case "pr$a"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"				
				errou="pr"				
			case "pr$0"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"			
				errou="pr"
			case "pr$1"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 8"				
				errou="pr"
			case "pr-b$1"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"				
				errou="pr"
			case "sim$a"
				wrt = "O campo Sim do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 3"				
				errou="sim"				
			case "sim$0"
				wrt = "O campo Sim do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="sim"
			case "sim$1"
				wrt = "O campo Sim do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 2"				
				errou="sim"
			case "bon$0"
				wrt = "O campo Bon do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="bon"
			case "bon"
				wrt = "O campo Bon do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="bon"
			case "rec$0"
				wrt = "O campo Rec do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="rec"
			case "rec"
				wrt = "O campo Rec do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 6"
				errou="rec"
		end select		
'================================================================================================================================		
	elseif ambiente_escola="jbarro" or ambiente_escola="testejbarro"  then
		SELECT CASE dados2
			case "f$0"
				wrt = "O campo faltas do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"
				errou="faltas"
			case "f$1"
				wrt = "O campo faltas do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor que 255"				
				errou="faltas"
			case "tr1$a"
				wrt = "O campo Tr1 n&atilde;o deve ser utilizado"				
				errou="tr1"				
			case "tr1$0"
				wrt = "O campo Tr1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="tr1"
			case "tr1$1"
				wrt = "O campo Tr1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="tr1"
			case "tr2$a"
				wrt = "O campo Tr2 n&atilde;o deve ser utilizado"				
				errou="tr2"						
			case "tr2$0"
				wrt = "O campo Tr2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="tr2"
			case "tr2$1"
				wrt = "O campo Tr2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"				
				errou="tr2"
			case "tr3$a"
				wrt = "O campo Tr3 n&atilde;o deve ser utilizado"				
				errou="tr3"						
			case "tr3$0"
				wrt = "O campo Tr3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="tr3"
			case "tr3$1"
				wrt = "O campo Tr3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="tr3"
			case "te1$a"
				wrt = "O campo Te1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 7"				
				errou="te1"				
			case "te1$0"
				wrt = "O campo Te1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="te1"
			case "te1$1"
				wrt = "O campo Te1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"				
				errou="te1"
			case "te2$a"
				wrt = "O campo Te2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 7"				
				errou="te2"					
			case "te2$0"
				wrt = "O campo Te2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"			
				errou="te2"
			case "te2$1"
				wrt = "O campo Te2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="te2"
			case "te3$a"
				wrt = "O campo Te3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 7"				
				errou="te3"					
			case "te3$0"
				wrt = "O campo Te3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"		
				errou="te3"
			case "te3$1"
				wrt = "O campo Te3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="te3"
			case "pr$a"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"				
				errou="pr"				
			case "pr$0"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"			
				errou="pr"
			case "pr$1"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 8"				
				errou="pr"
			case "pr-b$1"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"				
				errou="pr"
			case "sim$a"
				wrt = "O campo Sim do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 3"				
				errou="sim"				
			case "sim$0"
				wrt = "O campo Sim do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="sim"
			case "sim$1"
				wrt = "O campo Sim do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 2"				
				errou="sim"
			case "bon$0"
				wrt = "O campo Bon do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="bon"
			case "bon$1"
				wrt = "O campo Bon do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="bon"
			case "rec$0"
				wrt = "O campo Rec do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="rec"
			case "rec$1"
				wrt = "O campo Rec do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="rec"
		end select
		
		
'================================================================================================================================		









	elseif ambiente_escola="mraythe" or ambiente_escola="testemraythe"  then
		SELECT CASE dados2
			case "pt1"
				wrt = "O campo Peso deve ser númerico e menor que dez"
				errou="va_pt1"
			case "pt2"
				wrt = "O campo Peso deve ser númerico e menor que dez"
				errou="va_pt2"
			case "pt3"
				wrt = "O campo Peso deve ser númerico e menor que dez"
				errou="va_pt3"
			case "pt4"
				wrt = "O campo Peso deve ser númerico e menor que dez"
				errou="va_pt4"
			case "pt"
				wrt = "A soma dos Pesos deve ser menor que dez"
				errou="total_peso"																
			case "f$0"
				wrt = "O campo faltas do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"
				errou="faltas"
			case "f$1"
				wrt = "O campo faltas do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor que 255"				
				errou="faltas"
			case "t1-a$0"
				wrt = "O campo Tr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="tr1"
			case "t1-a$1"
				wrt = "O campo Tr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 1"
				errou="tr1"
			case "t1-a$2"
				wrt = "O campo Tr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="tr1"				
			case "tr1-b$0"
				wrt = "O campo Tr1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="tr1"
			case "tr1-b$1"
				wrt = "O campo Tr1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual ao respectivo peso"				
				errou="tr1"				
			case "t1-c$0"
				wrt = "O campo Tr1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="t1"
			case "t1-c$1"
				wrt = "O campo Tr1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual ao respectivo peso"
				errou="t1"				
			case "t2-a$0"
				wrt = "O campo S1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="tr2"
			case "t2-a$1"
				wrt = "O campo S1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 2"
				errou="tr2"
			case "t2-a$2"
				wrt = "Valor inv&aacute;lido para o campo S1 do(a) aluno(a) n&ordm;"&num_chamada_erro&". A soma dos simulados deve ser menor ou igual a 3."
				errou="tr2"	
			case "t2-a$3"
				wrt = "O campo S1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="tr2"							
			case "tr2-b$0"
				wrt = "O campo Tr2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="tr2"
			case "tr2-b$1"
				wrt = "O campo Tr2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual ao respectivo peso"				
				errou="tr2"				
			case "t2-c$0"
				wrt = "O campo Tr2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="t2"
			case "t2-c$1"
				wrt = "O campo Tr2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" eve ser menor ou igual ao respectivo peso"
				errou="t2"	
			case "t3-a$0"
				wrt = "O campo S2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="tr3"
			case "t3-a$1"
				wrt = "Valor inv&aacute;lido para o campo S2 do(a) aluno(a) n&ordm;"&num_chamada_erro&". A soma dos simulados deve ser menor ou igual a 4."
				errou="tr3"
			case "t3-a$2"
				wrt = "Valor inv&aacute;lido para o campo S2 do(a) aluno(a) n&ordm;"&num_chamada_erro&". A soma dos simulados deve ser menor ou igual a 3."
				errou="tr3"	
			case "t3-a$3"
				wrt = "O campo S2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10."
				errou="tr3"							
			case "tr3-b$0"
				wrt = "O campo Tr3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="tr3"
			case "tr3-b$1"
				wrt = "O campo Tr3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual ao respectivo peso"				
				errou="tr3"				
			case "t3-c$0"
				wrt = "O campo Tr3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="t3"
			case "t3-c$1"
				wrt = "O campo Tr3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual ao respectivo peso"
				errou="t3"	
			case "tr4-b$0"
				wrt = "O campo Tr4 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="tr4"
			case "tr4-b$1"
				wrt = "O campo Tr4 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual ao respectivo peso"				
				errou="tr4"				
			case "t4-c$0"
				wrt = "O campo Tr4 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="t4"
			case "t4-c$1"
				wrt = "O campo Tr4 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual ao respectivo peso"
				errou="t4"	
			case "t1-b$0"
				wrt = "O campo Te1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="t1"
			case "t1-b$1"
				wrt = "O campo Te1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="t1"	
			case "t2-b$0"
				wrt = "O campo Te2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="t2"
			case "t2-b$1"
				wrt = "O campo Te2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="t2"	
			case "t3-b$0"
				wrt = "O campo Te3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="t3"
			case "t3-b$1"
				wrt = "O campo Te3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="t3"	
			case "t4-b$0"
				wrt = "O campo Te4 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="t4"
			case "t4-b$1"
				wrt = "O campo Te4 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="t4"									
			case "p1-a$0"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="p1"
			case "p1-a$1"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 5"	
				errou="p1"
			case "p1-a$2"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"	
				errou="p1"	
			case "p1-a$3"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 6"	
				errou="p1"					
			case "p1-b$0"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 5"	
				errou="p1"
			case "p1-b$1"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"	
				errou="p1"	
			case "p1c$0"
				wrt = "O campo Te do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 5"	
				errou="p1"
			case "p1-c$1"
				wrt = "O campo Te do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"	
				errou="p1"					
			case "p2-a$0"
				wrt = "O campo Atv do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="p2"
			case "p2-a$1"
				wrt = "Valor inv&aacute;lido para o campo Atv do(a) aluno(a) n&ordm;"&num_chamada_erro&". A soma dde Pr com Atv deve ser menor ou igual a 10."
				errou="p2"
			case "p2-a$2"
				wrt = "O campo Atv do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="p2"				
			case "p2-c$0"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="p2"
			case "p2-c$1"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"	
				errou="p2"			
			case "bon$0"
				wrt = "O campo Ext do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="bon"
			case "bon$1"
				wrt = "O campo Ext do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"	
				errou="bon"				
			case "rec$0"
				wrt = "O campo Rec do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="rec"
			case "rec$1"
				wrt = "O campo Rec do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="rec"
		end select

'================================================================================================================================

	elseif ambiente_escola="stockler" or ambiente_escola="testestockler"  then
		SELECT CASE dados2
			case "f$0"
				wrt = "O campo faltas do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"
				errou="faltas"
			case "f$1"
				wrt = "O campo faltas do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor que 255"				
				errou="faltas"
			case "av1$0"
				wrt = "O campo AV1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="av1"
			case "av1$1"
				wrt = "O campo AV1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="av1"
			case "av2$0"
				wrt = "O campo AV2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="av2"
			case "av2$1"
				wrt = "O campo AV2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"				
				errou="av2"
			case "av3$0"
				wrt = "O campo AV3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="av3"
			case "av3$1"
				wrt = "O campo AV3 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="av3"
			case "pr$0"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="pr"
			case "pr$1"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"	
				errou="pr"
			case "at$0"
				wrt = "O campo At do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero"				
				errou="at"
			case "at$1"
				wrt = "O campo At do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"	
				errou="at"				
			case "rec$0"
				wrt = "O campo Rec do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="rec"
			case "rec$1"
				wrt = "O campo Rec do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 10"
				errou="rec"
		end select

'================================================================================================================================			
		
	elseif ambiente_escola="vitoria" or ambiente_escola="testevitoria"  then
		SELECT CASE dados2
			case "f$0"
				wrt = "O campo faltas do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"
				errou="faltas"
			case "f$1"
				wrt = "O campo faltas do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor que 255"				
				errou="faltas"
			case "te1$0"
				wrt = "O campo Te1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="te1"
			case "te1$1"
				wrt = "O campo Te1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 100"
				errou="te1"
			case "te2$0"
				wrt = "O campo Te2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="te2"
			case "te2$1"
				wrt = "O campo Te2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 100"				
				errou="te2"
			case "tr1$0"
				wrt = "O campo Tr1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="tr1"
			case "tr1$1"
				wrt = "O campo Tr1 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 100"
				errou="tr1"
			case "tr2$0"
				wrt = "O campo Tr2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="tr2"
			case "tr2$1"
				wrt = "O campo Tr2 do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 100"				
				errou="tr2"
			case "am$0"
				wrt = "O campo Am do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="am"
			case "am$1"
				wrt = "O campo Am do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 100"				
				errou="am"
			case "ab$0"
				wrt = "O campo Ab do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="ab"
			case "ab$1"
				wrt = "O campo Ab do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 100"
				errou="ab"
			case "pr$0"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="pr"
			case "pr$1"
				wrt = "O campo Pr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 100"	
				errou="pr"
			case "sim$0"
				wrt = "O campo Sim do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="sim"
			case "sim$1"
				wrt = "O campo Sim do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 60"				
				errou="sim"				
			case "cf$0"
				wrt = "O campo Cf do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="cf"
			case "cf$1"
				wrt = "O campo Cf do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 40"		
				errou="cf"
			case "cf$1c"
				wrt = "O campo Cf do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 100"		
				errou="cf"				
			case "bon$0"
				wrt = "O campo Bon do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="bon"
			case "bon$1"
				wrt = "O campo Bon do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 100"
				errou="bon"
			case "bon$2"
				wrt = "A soma de M1 com o campo Bon do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 100"
				errou="bon"				
			case "rec$0"
				wrt = "O campo Rec do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="rec"
			case "recb$1"
				wrt = "O campo Rec do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 70"
				errou="rec"
			case "rec$1"
				wrt = "O campo Rec do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 100"
				errou="rec"
			case "cfr$0"
				wrt = "O campo Cfr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve conter um n&uacute;mero inteiro"				
				errou="cfr"
			case "cfr$1"
				wrt = "O campo Cfr do(a) aluno(a) n&ordm;"&num_chamada_erro&" deve ser menor ou igual a 30"		
				errou="cfr"				
		end select
		
		
'================================================================================================================================			
																
	end if					
'Mensagens Gerais de 0 a 100
case 0
wrt = "Escolha uma das op&ccedil;&otilde;es abaixo"

case 1
wrt = "Selecione uma unidade, um curso, uma etapa e uma turma. "

case 2
wrt = "Selecione uma etapa e uma turma."

case 3
wrt = "Selecione uma etapa, uma turma, um per&iacute;odo e uma avalia&ccedil;&atilde;o."

case 4
wrt = "Para consultar &eacute; necess&aacute;rio selecionar uma etapa!"

case 5
wrt = "Esta fun&ccedil;&atilde;o permite voc&ecirc; fazer contato com a equipe t&eacute;cnica que realiza a manuten&ccedil;&atilde;o do sistema Web Diretor. Utilize sempre que poss&iacute;vel este canal para nos transmitir alguma informa&ccedil;&atilde;o relevante sobre o funcionamento desse produto. Obrigado pela sua aten&ccedil;&atilde;o!"

case 6
wrt = "Mensagem enviada."

case 7
wrt = "Escolha um novo usu&aacute;rio."

case 8
wrt = "Escolha uma nova senha."

case 9
wrt = "Usu&aacute;rio alterado com sucesso."

case 10
wrt = "Senha alterada com sucesso."

case 11
wrt = "Selecione uma disciplina e um per&iacute;odo."

case 12
wrt = "E-mail alterado com sucesso."

case 13
wrt = "Usu&aacute;rio j&aacute; existe!"

case 14
wrt = "Digite seu novo endere&ccedil;o de correio eletrônico"

case 15
wrt = "Endere&ccedil;o de correio eletrônico j&aacute; existe!"

case 16
wrt = "Selecione uma etapa, uma turma e um per&iacute;odo."

case 17
wrt = "Selecione uma etapa e um per&iacute;odo."

case 18
wrt = "Gr&aacute;fico comparativo."

case 19
wrt = "Selecione uma etapa, uma disciplina e um per&iacute;odo."

case 20
wrt = "Selecione uma etapa"

case 21
wrt = "Clique na op&ccedil;&atilde;o desejada"

case 22
wrt = "Confirma a reinicializa&ccedil;&atilde;o da senha do usu&aacute;rio abaixo?"

case 23
wrt = "Senha reinicializada com sucesso"

case 24
wrt = "Usu&aacute;rio "&situacao&" com sucesso"





'Web Fam&iacute;lia de 100 a 199
case 100
wrt = "Selecione o tipo de documento e os arquivos que deseja disponibilizar para upload"

case 101
wrt = "Arquivo(s) "&Session("arquivos") &" enviado(s) com sucesso! Total de Bytes enviados:"&Session("upl_total")

case 102
wrt = "Selecione pelo menos um arquivo"

case 103
wrt = "Preencha os dados abaixo para associar um documento"

case 104
wrt = "Associa&ccedil;&atilde;o realizada com Sucesso"

case 105
wrt = "Preencha os dados abaixo para incluir uma not&iacute;cia"

case 106
wrt = "Not&iacute;cia inclu&iacute;da com sucesso"

case 107
wrt = "Confirma a exclus&atilde;o do(s) documento(s) abaixo?"

case 108
wrt = "Documento(s) exclu&iacute;do(s) com sucesso"

case 109
wrt = "Confirma a exclus&atilde;o do(s) arquivo(s) abaixo?"

case 110
wrt = "Arquivo(s) exclu&iacute;do(s) com sucesso"

case 111
wrt = "Selecione o tipo de documento"

case 112
wrt = "Confirma a exclus&atilde;o da(s) not&iacute;cia(s) abaixo?"

case 113
wrt = "Not&iacute;cia(s) exclu&iacute;da(s) com sucesso"

case 114
wrt = "Confirma a exclus&atilde;o do(s) evento(s) abaixo?"

case 115
wrt = "Evento(s) exclu&iacute;do(s) com sucesso"

case 116
wrt = "Preencha os dados abaixo para incluir um evento"

case 117
wrt = "Evento inclu&iacute;do com sucesso"

case 118
wrt = "Para consultar os dados do usu&aacute;rio digite o c&oacute;digo ou Nome e clique no bot&atilde;o Procurar."

case 119
wrt = "Escolha um usu&aacute;rio para consultar o cadastro."

case 120
wrt = "Verifique os dados do usu&aacute;rio."

case 121
wrt = "N&atilde;o  foi encontrado nenhum usu&aacute;rio com este c&oacute;digo."

' erro na busca por nome
case 122
wrt = "N&atilde;o  foi encontrado nenhum usu&aacute;rio com este nome."


'alunos de 300 a 499
case 300
wrt = "Para consultar os dados do Aluno digite a matr&iacute;cula ou Nome e clique no bot&atilde;o Procurar."

' listagem de alunos

case 301
wrt = "Escolha um Aluno para consultar o cadastro."

case 302
wrt = "Verifique os dados do Aluno."

case 303
wrt = "N&atilde;o  foi encontrado nenhum Aluno com este c&oacute;digo."

' erro na busca por nome
case 304
wrt = "N&atilde;o  foi encontrado nenhum Aluno com este nome."

case 305
wrt = "Lista de alunos associados a turma abaixo."

case 306
wrt = "Verifique os dados dos familiares."

case 307
wrt = "Selecione uma unidade e um m&ecirc;s."

case 308
wrt = "Comparar Turma por M&eacute;dia Geral."

case 309
wrt = "Verifique os dados do Aluno e escolha uma disciplina e um per&iacute;odo."

case 310
wrt = "Escolha os crit&eacute;rios para pesquisar as ocorr&ecirc;ncias do aluno e clique no bot&atilde;o prosseguir."

case 311
wrt = "Confirma a exclus&atilde;o dessa(s) disciplina(s)."


case 312
wrt = "Ocorr&ecirc;ncia inclu&iacute;da com sucesso!"

case 313
wrt = "Ocorr&ecirc;ncia alterada com sucesso!"

case 314
wrt = "Ocorr&ecirc;ncia exclu&iacute;da com sucesso!"

case 315
wrt = "Preencha os dados abaixo e clique no bot&atilde;o Confirmar para Incluir uma nova ocorr&ecirc;ncia."

case 316
wrt = "Preencha os dados abaixo e clique no bot&atilde;o Confirmar para atualizar esta ocorr&ecirc;ncia."

case 317
wrt = "Selecione uma situa&ccedil;&atilde;o para o aluno e escreva o motivo da inativa&ccedil;&atilde;o."



'web secretaria 500 a 599
case 400
wrt = "Para consultar os dados do Aluno digite a matr&iacute;cula ou Nome e clique no bot&atilde;o Procurar. Caso o aluno N&atilde;o  esteja cadastrado no sistema clique <a href='../../../cad/man/aal/cadastra.asp?nvg=WS-CA-MA-AAL' class='avisos'>aqui</a>."

case 401
wrt = "matr&iacute;cula efetuada com sucesso!"

case 402
wrt = "Preencha os campos abaixo."

case 403
wrt = "Aluno j&aacute; matriculado para este ano letivo. matr&iacute;culas para o pr&oacute;ximo Ano Letivo est&atilde;o fechadas!"

case 404
wrt = "Para alterar os dados do Aluno digite a matr&iacute;cula ou Nome e clique no bot&atilde;o Procurar. Caso o aluno N&atilde;o  esteja cadastrado no sistema clique <a href='../../../cad/man/aal/cadastra.asp?nvg=WS-CA-MA-AAL' class='avisos'>aqui</a>."

case 405
dados=dados

separa=split(dados,"#sep#")
ordem_familiares=separa(0)
qtd_tipo_familiares=separa(1)
cod_familiar=separa(2)
cod_vinculado=separa(3)
cod_aluno=separa(4)
wrt1 ="<input name='ordem' type='hidden' value='"&ordem_familiares&"'>"
'wrt2 ="<input name='cod_prim' type='hidden' value='"&cod_familiar_prim&"'>"
wrt2 ="<input name='qtd' type='hidden' value='"&qtd_tipo_familiares&"'>"
wrt3 ="<input name='foco' type='hidden' value='"&cod_familiar&"'>"
wrt4 ="<input name='cod_vinculado' type='hidden' value='"&cod_vinculado&"'>"
wrt5 ="<input name='cod_al' type='hidden' value='"&cod_aluno&"'>"
wrt6 =Server.URLEncode("Confirma a exclus&atilde;o desse familiar?")

wrt = wrt1&wrt2&wrt3&wrt4&wrt5&wrt6&"<br><input type='button' name='Submit2' value='Sim' onClick='ExcluiFamiliares(ordem.value,qtd.value,foco.value,cod_al.value)' class='botao_prosseguir_sim' >&nbsp;&nbsp;&nbsp;<input type='button' name='Submit2' value='"&Server.URLEncode("N&atilde;o ")&"' onClick='recuperarFamiliares(ordem.value,qtd.value,foco.value,cod_vinculado.value,cod_al.value)' class='botao_prosseguir_nao' >"

case 406
dados=dados

separa=split(dados,"#sep#")
ordem_familiares=separa(0)
qtd_tipo_familiares=separa(1)
cod_familiar=separa(2)
cod_vinculado=separa(3)
cod_aluno=separa(4)
'cod_nome = Split(ordem_familiares, "!!")
'cod_familiar_prim=cod_nome(0)
wrt1 ="<input name='ordem' type='hidden' value='"&ordem_familiares&"'>"
'wrt2 ="<input name='cod_prim' type='hidden' value='"&cod_familiar_prim&"'>"
wrt2 ="<input name='qtd' type='hidden' value='"&qtd_tipo_familiares&"'>"
wrt3 ="<input name='foco' type='hidden' value='"&cod_familiar&"'>"
wrt4 ="<input name='cod_vinculado' type='hidden' value='"&cod_vinculado&"'>"
wrt5 ="<input name='cod_al' type='hidden' value='"&cod_aluno&"'>"
wrt6 =Server.URLEncode("O CPF Digitado possui dados cadastrados. Deseja aproveitar esses dados?")

wrt = wrt1&wrt2&wrt3&wrt4&wrt5&wrt6&"<br><input type='button' name='Submit2' value='Sim' onClick='recuperarFamiliares(ordem.value,qtd.value,foco.value,cod_vinculado.value,cod_al.value)' class='botao_prosseguir_sim' >&nbsp;&nbsp;&nbsp;<input type='button' name='Submit2' value='"&Server.URLEncode("N&atilde;o ")&"' onClick='ExcluiFamiliares(ordem.value,qtd.value,foco.value,cod_al.value)' class='botao_prosseguir_nao' >"

case 407
wrt = "Deve ser selecionado um respons&aacute;vel financeiro para o aluno!"

case 408
wrt = "Deve ser selecionado um respons&aacute;vel pedag&oacute;gico para o aluno!"

case 409
wrt = "&eacute; obrigat&oacute;rio o preenchimento dos campos: Nome, Telefones de Contato e Endere&ccedil;o residencial para o respons&aacute;vel financeiro!"

case 410
wrt = "&eacute; obrigat&oacute;rio o preenchimento dos campos: Nome, Telefones de Contato e Endere&ccedil;o residencial para o respons&aacute;vel pedag&oacute;gico!"

case 411
wrt = "Ao se confirmar o cadastro desse aluno, esse n&uacute;mero de matr&iacute;cula N&atilde;o  poder&aacute; mais ser utilizado!"

case 412
wrt = "Cadastro efetuado com sucesso! Inclua todos os dados necess&aacute;rios."

case 413
wrt = "Selecione uma nova combina&ccedil;&atilde;o de Unidade, Curso, Etapa, Turma e n&uacute;mero de chamada para o aluno."

case 414
wrt = "Selecione um m&eacute;todo para enturmar os alunos em situa&ccedil;&atilde;o de pr&eacute;-matr&iacute;cula."

case 415
wrt = "N&atilde;o  existem alunos em situa&ccedil;&atilde;o de pr&eacute;-matr&iacute;cula."

case 416
wrt = "Somente &eacute; poss&iacute;vel remanejar alunos com situa&ccedil;&atilde;o igual a 'Cursando'."


'professores de 600 a 799

case 600
wrt =  "Os Professores em vermelho est&atilde;o inativos. A mensagem 'N&atilde;o  cadastrado' indica que N&atilde;o  existe professor associado &agrave;quela disciplina naquela turma"
wrt = wrt &"<br>A mensagem 'nome em branco' indica que o nome do professor N&atilde;o  est&aacute; registrado no cadastro. Para bloquear a planilha clique na letra 'N' do per&iacute;odo escolhido"

case 601
wrt = "Confirma o " 
if opt="blq" then
wrt= wrt &"BLOQUEIO"
else
wrt= wrt &"DESBLOQUEIO"
end if
wrt= wrt &" das notas do trimestre "&periodo&" de "&no_materia&", Unidade:"&no_unidade&" - "&no_etapa&" do "&no_curso&" Turma "&turma&""

case 602
if orig=01 then
act= "bloqueada"
elseif orig=02 then
act= "desbloqueada"
end if

wrt = "Planilha "&act&" com sucesso!"

case 603
wrt = "Avalia&ccedil;&otilde;es N&atilde;o  lan&ccedil;adas"

case 604
wrt = "Para consultar a Grade de aulas digite o C&oacute;digo ou Nome de um Professor e clique no bot&atilde;o Procurar."
wrt = wrt &"<br>Se preferir obter uma lista completa de TODOS os professores clique <a href='index.asp?opt=listall&nvg="&nvg&"' class='avisos'>aqui</a>"

case 605
wrt = "N&atilde;o  foi encontrado nenhum professor com este c&oacute;digo."

case 606
wrt = "Escolha um professor para consultar a Grade de Aulas. Os Professores em vermelho est&atilde;o inativos."

case 607
wrt = "Para atualizar os dados do Professor digite o C&oacute;digo ou Nome e clique no bot&atilde;o Procurar."
wrt = wrt &"Se preferir adicionar um NOVO professor clique <a href='altera.asp?ori=02&nvg="&nvg&"' class='avisos'>aqui</a>."
wrt = wrt &"<BR>Se preferir obter uma lista completa de TODOS os professores clique <a href='index.asp?opt=listall&nvg="&nvg&"' class='avisos'>aqui</a>"

case 608
wrt = "Confirme o professor para consultar a Grade de Aulas."

case 609
wrt = "O per&iacute;odo relacionado pela letra 'S' indica que a planilha est&aacute; Bloqueada e 'N' que est&aacute; Desbloqueada."

case 610
wrt = "N&atilde;o  foi encontrado nenhum professor com este c&oacute;digo."

case 611
wrt = "N&atilde;o  foi encontrado nenhum professor com este nome."

case 612
wrt = "Escolha um professor para atualizar o cadastro. Os Professores em vermelho est&atilde;o inativos."

case 613
wrt = "Confirme se &eacute; o professor correto para atualizar o cadastro."

case 614
wrt = "Preencha cuidadosamente os dados do Professor e click no bot&atilde;o CONFIRMAR para atualizar o cadastro"

case 615
wrt = "Professor c&oacute;digo "&cod_cons&" e usu&aacute;rio "&escola&co_usr_prof&" inclu&iacute;do com sucesso!"

case 616
wrt = "Dados do Professor c&oacute;digo "&cod_cons&" alterados com sucesso!"

case 617
wrt = "Selecione a Data e a Hora as quais voc&ecirc; deseja iniciar o monitoramento de notas e clique em iniciar."

case 618
mes_mnl=mes_mnl*1
min_mnl=min_mnl*1
			  if mes_mnl< 10 then
			  mes_wrt="0"&mes_mnl
			  else
			  mes_wrt=mes_mnl					  
			  end if 
					  
			  if min_mnl< 10 then
			  min_wrt="0"&min_mnl
			  else
			  min_wrt=min_mnl					  
			  end if 
wrt = "Inicio da monitora&ccedil;&atilde;o a partir do dia "&dia_mnl&"/"&mes_wrt&"/"&ano_mnl&" as "&hora_mnl&":"&min_wrt&" Dados atualizados a cada minuto."

case 619
wrt = "N&atilde;o foram encontradas turmas cadastradas para voc&ecirc;. Entre em contato com o seu coordenador."


case 620
if errou="pv1" or errou="pv2" or errou="pv3" or errou="pv4" or errou="pv5" or errou="pv6" Then
wrt = "Valor inv&aacute;lido para o campo  "&errado
elseif errou="sp" Then
wrt = "Soma dos Pesos maior que 10"
elseif errou="pt" Then
wrt = "Um dos pesos tem valor inv&aacute;lido"
elseif errou="pr1pr2" Then
wrt = "Soma das Pr's maior que 10"
else
wrt = "Valor inv&aacute;lido para o campo  "&errado&"  do n&uacute;mero de chamada <b>"&errante&"</b>"
end if

' erro na busca por c&oacute;digo
case 621
wrt = "Voc&ecirc; est&aacute; " 
if opt="cln" then
wrt= wrt &"comunicando"
else
wrt= wrt &"lan&ccedil;ando"
end if


		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set RSpr = Server.CreateObject("ADODB.Recordset")
		SQLpr = "SELECT * FROM TB_Periodo where NU_Periodo = "&periodo
		RSpr.Open SQLpr, CON0

no_periodo=RSpr("NO_Periodo")

wrt= wrt &" notas de "&no_periodo&" de "&no_materia&", Unidade:"&no_unidades&" - "&no_serie&" do "&no_grau&" Turma "&turma&""

case 622
wrt = "Notas lan&ccedil;adas com sucesso."

case 623
wrt = "Comunicado efetuado!"

case 624
wrt = "Estas notas j&aacute; foram lan&ccedil;adas.Para alter&aacute;-las pe&ccedil;a autoriza&ccedil;&atilde;o ao coordenador"

case 625
wrt = "Escolha um Coordenador para consultar os Professores sob sua coordena&ccedil;&atilde;o."

case 626
wrt = "Os Professores em vermelho est&atilde;o inativos. A mensagem 'N&atilde;o  cadastrado'indica que N&atilde;o  existe professor associado &agrave;quela disciplina naquela turma"
wrt = wrt &"<br>A mensagem 'nome em branco' indica que o nome do professor N&atilde;o  est&aacute; registrado no cadastro"

case 627
wrt = "Para excluir, selecione uma ou mais disciplinas e clique em excluir.<br>Para incluir uma nova disciplina na Grade de Aulas, selecione uma unidade e um curso."

case 628
wrt = "Disciplina inclu&iacute;da com sucesso"

case 629
wrt = "Disciplina exclu&iacute;da com sucesso"

case 630
wrt = "N&atilde;o  &eacute; poss&iacute;vel marcar uma disciplina na Grade de Aulas e selecionar uma unidade e um curso ao mesmo tempo.<br>Por favor selecione somente disciplina(s) para excluir ou selecione uma unidade para incluir uma nova disciplina na Grade de Aulas"

case 631
wrt = "Selecione uma disciplina, um modelo e um coordenador."

case 632
wrt = "Para atualizar &eacute; necess&aacute;rio selecionar uma disciplina,um modelo e um coordenador"

case 633
wrt = "Verifique os dados preenchidos e clique no bot&atilde;o Confirmar para continuar a inclus&atilde;o ou no bot&atilde;o Alterar para voltar e modificar algum dado."


case 634
wrt = "Verifique as disciplinas selecionadas e clique no bot&atilde;o confirmar para Excluir ou no bot&atilde;o Cancelar para voltar e modificar algum dado."

case 635
wrt = "Professores que N&atilde;o  comunicaram."

case 636
wrt = "Para imprimir clique <a class='avisos' href='#' onClick=MM_openBrWindow('imprime.asp?or=01&obr="&obr&"&p=p','','status=yes,menubar=yes,scrollbars=yes,resizable=yes,width=1030,height=500,top=50,left=50')>aqui</a>."

case 637
wrt = "Escolha um professor e um per&iacute;odo."

case 638
wrt =  "Os Professores em vermelho est&atilde;o inativos. A mensagem 'N&atilde;o  cadastrado' indica que N&atilde;o  existe professor associado &agrave;quela disciplina naquela turma"
wrt = wrt &"<br>A mensagem 'nome em branco' indica que o nome do professor N&atilde;o  est&aacute; registrado no cadastro. Clique no nome da disciplina para ver o mapa de resultado."

case 639
wrt = "Arquivo "& fl &" enviado com sucesso."

case 640
wrt = "Aten&ccedil;&atilde;o! Estas notas j&aacute; foram lan&ccedil;adas pelo professor."

case 641
wrt = "Inclua as faltas no per&iacute;odo desejado"

case 642
wrt = "Faltas lan&ccedil;adas com sucesso"

case 643
wrt = "Para atualizar os dados do Professor digite o C&oacute;digo ou Nome e clique no bot&atilde;o Procurar."
wrt = wrt &"<BR>Se preferir obter uma lista completa de TODOS os professores clique <a href='index.asp?opt=listall&nvg="&nvg&"' class='avisos'>aqui</a>"

case 644
wrt = "&eacute; necess&aacute;rio escolher pelo menos uma unidade"

case 645
wrt = "Imprimir <a class='avisos' href='#' onClick=MM_openBrWindow('imprime.asp?obr="&obr&"&p=p','','status=yes,menubar=yes,scrollbars=yes,resizable=yes,width=1030,height=500,top=50,left=50')>html</a> / <a class='avisos' href='../../../../relatorios/swd015.asp?obr="&obr&"'>pdf</a>."


'Mensagens de web Tesouraria de 700 a 999
case 700
wrt = "Selecione o local da base de dados Posicao.mdb"

case 701
wrt = "&eacute; necess&aacute;rio selecionar o local da base de dados Posicao.mdb"

case 702
wrt = "O nome do arquivo deve obrigatoriamente ser Posicao.mdb"

case 703
wrt = "Base de dados financeiros atualizada com sucesso!"


'Mensagens de sistema de 9700 a 9999
case 9700
wrt = "Acesso N&atilde;o  permitido a esta fun&ccedil;&atilde;o!"

case 9701
wrt = "Acesso permitido somente para consulta!"

case 9702
wrt = "Para imprimir clique <a class='avisos' href='#' onClick=MM_openBrWindow('imprime.asp?or=01&obr="&obr&"&p=p','','status=yes,menubar=yes,scrollbars=yes,resizable=yes,width=1030,height=500,top=50,left=50')>aqui</a>."

case 9703
wrt = "Aten&ccedil;&atilde;o! Ano Letivo est&aacute; Finalizado. As fun&ccedil;&otilde;es s&oacute; poder&atilde;o ser consultadas!<a href=../inicio.asp><img src=../img/ok.gif align=absbottom></a>"

case 9704
wrt = "Selecione as op&ccedil;&otilde;es desejadas."

case 9705
wrt = "Dados alterados com sucesso!"

case 9706
wrt = "Selecione os par&acirc;metros desejados"

case 9707
wrt = "Resultado encontrado de acordo com par&acirc;metros informados"

case 9708
wrt = "Altere os dados necess&aacute;rios"

case 9709
wrt = "Dados alterados com sucesso"

case 9710
wrt = "ERRO!"

case 9711
wrt = "Digite a matr&iacute;cula ou o nome do aluno"


end select




SELECT CASE tab


' primeira tela
case "inf"

%>
<table width="1000" height="52" border="3" align="center" cellpadding="0" cellspacing="0" bordercolor="#EEEEEE" class="aviso1">
  <tr> 
            
    <td height="46"> <div align="center"> 
      <%SELECT CASE nivel
				case 0%>
      <img src="img/atencao.gif" width="23" height="25" align="absmiddle"> 
      <%case 1%>
      <img src="../img/atencao.gif" width="23" height="25" align="absmiddle"> 
      <%		case 2%>
      <img src="../../img/atencao.gif" width="23" height="25" align="absmiddle"> 
      <%		case 3%>
      <img src="../../../img/atencao.gif" width="23" height="25" align="absmiddle"> 
      <%		case 4%>
      <img src="../../../../img/atencao.gif" width="23" height="25" align="absmiddle"> 
      <%end select%>
		  <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>
          <%response.Write(wrt)%>
          </strong></font></div>
                </div></td>
          </tr>
        </table>
		
<%
' erro
case "err"
%>
<table width="1000" height="30" border="3" align="center" cellpadding="0" cellspacing="0" bordercolor="#EEEEEE" bgcolor="#FFE8E8" class="aviso2">
  <tr> 
            <td> <div align="center"> 
                <p>
		<%SELECT CASE nivel
				case 0%>
				<img src="img/pare.gif" width="28" height="25" align="absmiddle">
		<%case 1%>
				<img src="../img/pare.gif" width="28" height="25" align="absmiddle">
		<%case 2%>
				<img src="../../img/pare.gif" width="28" height="25" align="absmiddle">
		<%case 3%>
				<img src="../../../img/pare.gif" width="28" height="25" align="absmiddle">
		<%case 4%>
				<img src="../../../../img/pare.gif" width="28" height="25" align="absmiddle">												
		<%end select%>
                <font color="#CC0000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%response.Write(wrt)%></strong></font> 
                </p>
              </div></td>
          </tr>
        </table>
<%
' inclus&atilde;o / altera&ccedil;&atilde;o de dados
case "ok"
%>
<table width="1000" height="30" border="3" align="center" cellpadding="0" cellspacing="0" bordercolor="#EEEEEE" bgcolor="#F2F9EE">
  <tr> 
            <td> <div align="center"> 
        <p>
		<%SELECT CASE nivel
						case 0%>
				<img src="img/atencao2.gif" width="23" height="25" align="absmiddle">
		<%case 1%>
		<img src="../img/atencao2.gif" width="23" height="25" align="absmiddle"> 
		<%case 2%>
				<img src="../../img/atencao2.gif" width="23" height="25" align="absmiddle">
		<%case 3%>
				<img src="../../../img/atencao2.gif" width="23" height="25" align="absmiddle">
		<%case 4%>
				<img src="../../../../img/atencao2.gif" width="23" height="25" align="absmiddle">		
		<%end select%>		
          <font color="#CC0000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
          <%response.Write(wrt)%>
          </strong></font> </p>
              </div></td>
          </tr>
        </table>
<%
end select

End Function
%>