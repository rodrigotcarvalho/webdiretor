<!--#include file="../../global/funcoes_diversas.asp"-->
<%
Function calcular_nota(tipo_calculo,CAMINHOn,tb,nu_matricula,mat_princ,co_materia,periodo)
pasta=split(CAMINHOn, "\")
escola=pasta(4)
if escola="boechat" or escola="testeboechat" then

		Set CON_N = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIR3
		
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
		Set RS3 = CON_N.Execute(SQL_N)			

	if tipo_calculo="CALC1" then
	'Cálculo do NPA	

		IF RS3.EOF then
			npa="&nbsp;"
		else
			p_av1=RS3("VA_MAX_AV1")
			p_av2=RS3("VA_MAX_AV2")
			p_av3=RS3("VA_MAX_AV3")
			p_av4=RS3("VA_MAX_AV4")
			p_av5=RS3("VA_MAX_AV5")	
			va_av1=RS3("VA_AV1")
			va_av2=RS3("VA_AV2")	
			va_av3=RS3("VA_AV3")
			va_av4=RS3("VA_AV4")	
			va_av5=RS3("VA_AV5")		
	
			IF p_av1="" or isnull(p_av1) then
				s_av1=0	
				soma_av1=0				
			ELSE
				if va_av1="" or isnull(va_av1) then
					s_av1=0	
					soma_av1=0	
				else
					s_av1=1		
					soma_av1=va_av1					
				end if
			END IF	
			
			IF p_av2="" or isnull(p_av2) then
				s_av2=0	
				soma_av2=0				
			ELSE		
				if va_av2="" or isnull(va_av2) then
					s_av2=0	
					soma_av2=0	
				else
					s_av2=1		
					soma_av2=va_av2				
				end if
			END IF	
			
			IF p_av3="" or isnull(p_av3) then
				s_av3=0	
				soma_av3=0			
			ELSE								
				if va_av3="" or isnull(va_av3) then
					s_av3=0	
					soma_av3=0
				else
					s_av3=1		
					soma_av3=va_av3					
				end if
			END IF

			IF p_av4="" or isnull(p_av4) then
				s_av4=0	
				soma_av4=0		
			ELSE			
				if va_av4="" or isnull(va_av4) then	
					s_av4=0	
					soma_av4=0
				else
					s_av4=1	
					soma_av4=va_av4						
				end if	
			END IF	


			IF p_av5="" or isnull(p_av5) then
				s_av4=0	
				soma_av4=0		
			ELSE			
				if va_av5="" or isnull(va_av5) then		
					s_av5=0
					soma_av5=0	
				else
					s_av5=1		
					soma_av5=va_av5						
				end if
			END IF	
			
			if s_av1=1 or s_av2=1 or s_av3=1 or s_av4=1 or s_av5=1 then
				npa=soma_av1+soma_av2+soma_av3+soma_av4+soma_av5
				npa = arredonda(npa,"mat",0,0)
			else
				npa="&nbsp;"
			end if		
		END IF	
			resultado_calculo=npa
	else
		resultado_calculo="ERRO-1"
	end if	
else
	resultado_calculo="ERRO-2"	
end if	
calcular_nota=resultado_calculo
End Function	
%>