<!--#include file="caminhos.asp" -->
<%

	Set CONG = Server.CreateObject("ADODB.Connection") 
	ABRIRG = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONG.Open ABRIRG
	
Function Verifica_Conselho_Classe(cod_cons, disciplina, tipo_media, outro)

		Set RSVCC = Server.CreateObject("ADODB.Recordset")
		SQLVCC = "Select * from TB_Coc WHERE CO_Matricula = "& cod_cons &" AND  CO_Materia = '"& disciplina &"'"
		'response.write(SQLVCC&"<BR>")
					
		Set RSVCC = CONG.Execute(SQLVCC)
			
		if RSVCC.EOF then
			result = "N"	
		else		
		'response.write(RSVCC("STatus1")&"-"&RSVCC("STatus2")&"-"&RSVCC("STatus3")&"<BR>")
			
			if tipo_media = "MA" then'
				result = RSVCC("STatus1")	
			elseif tipo_media = "MF" then
				result = RSVCC("STatus2")						
			elseif tipo_media = "RF" or tipo_media = "T" then
				if RSVCC("STatus1")= "APC" or RSVCC("STatus2")= "APC" or RSVCC("STatus3")= "APC" then
					result = "APC"
				end if				
			end if				
		end if	
		if isnull(result) or result="" then
			result = "N"
		end if
		Verifica_Conselho_Classe = result
end function

'=============================================================================================================================
Function BonusMediaAnual(matricula, disciplina)


	Set RSBMA = Server.CreateObject("ADODB.Recordset")
	SQLBMA = "Select * from TB_Bonus_Media_Anual WHERE CO_Matricula = "& matricula
	Set RSBMA = CONG.Execute(SQLBMA)
	
	if RSBMA.eof then
		bonusBd = 0
	else
		bonusBd = RSBMA("bonus")
	end if	
	
	BonusMediaAnual	 = bonusBd

end function
Function TrataBonusMediaAnual(matricula, disciplina, tipo)
'tipo = "D" Decimal

	bonusBd = BonusMediaAnual(matricula, disciplina)
	
	'if tipo="D" then		
	'	TrataBonusMediaAnual = bonusBd/10	
	'else
		TrataBonusMediaAnual = bonusBd		
	'end if	

end function

Function AcrescentaBonusMediaAnual(matricula, disciplina, mediaAnual)

	acrescenta = TrataBonusMediaAnual(matricula, disciplina, "D")
	
	if isnumeric(mediaAnual) and session("ano_letivo") > 2012 then
		mediaAnual = mediaAnual*1
		acrescenta = acrescenta*1
		somaBonus = mediaAnual+acrescenta
		if somaBonus>100 then
			AcrescentaBonusMediaAnual = 100
		else
			AcrescentaBonusMediaAnual = somaBonus
		end if
	else
		AcrescentaBonusMediaAnual = mediaAnual
	end if	
end function

%>