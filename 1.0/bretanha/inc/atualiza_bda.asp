<!--#include file="caminhos.asp"-->
<%
Function Grava_BDA(p_nu_matricula, p_co_materia_pr, p_co_materia, p_periodo, p_va_faltas, p_va_pt, p_va_pp, p_va_t1, p_va_t2, p_va_t3, p_va_p1, p_va_p2, p_va_p3, p_va_bon, p_va_rec, p_data, p_horario, p_co_usr,p_todas_mt_subs_lancadas, p_todas_mp_subs_lancadas)

grava="ok"

nu_matricula = p_nu_matricula
co_materia_pr = p_co_materia_pr
co_materia = p_co_materia
periodo = p_periodo
va_faltas = p_va_faltas
va_pt = p_va_pt
va_pp = p_va_pp
va_t1 = p_va_t1
va_t2 = p_va_t2
va_t3 = p_va_t3
va_p1 = p_va_p1
va_p2 = p_va_p2
va_p3 = p_va_p3
va_bon = p_va_bon
va_rec = p_va_rec
data = p_data
horario = p_horario
co_usr = p_co_usr



		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_na & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

'////////////////////////////////////////////////////////////////
'FALTAS
	if va_faltas="" or isnull(va_faltas) then
		va_faltas=0			
	else
		teste_va_faltas = isnumeric(va_faltas)
		if teste_va_faltas= true then					
			va_faltas=va_faltas*1
			if va_faltas =<255 then
						IF Int(va_faltas)=va_faltas THEN
						va_faltas=va_faltas*1
						else	
							if  fail = 1 then
								grava = "no"
							else
								fail = 1 
								erro = "f"
								url = nu_matricula&"_"&va_faltas&"_"&erro
								grava = "no"
							end if
						end if
			else	
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "f"
					url = nu_matricula&"_"&va_faltas&"_"&erro
					grava = "no"
				end if
			end if		
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "f"
				url = nu_matricula&"_"&va_faltas&"_"&erro
				grava = "no"
			end if
		end if
	end if
'PESOS
	if va_pt="" or isnull(va_pt) then
		va_pt=1		
		s_va_pt=0
		soma_va_pt=0		
	else
		teste_va_pt = isnumeric(va_pt)
		if teste_va_pt= true then					
		va_pt=va_pt*1			
					if va_pt =<100 then
						IF Int(va_pt)=va_pt THEN
							s_va_pt=1
							soma_va_pt=va_pt						
						ELSE	
							if  fail = 1 then
								grava = "no"
							else
								fail = 1 
								erro = "pt"
								matric_Erro=i
								url = nu_matricula&"_"&va_pt&"_"&erro
								grava = "no"
							end if					
						end if																				
					else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "pt"
								matric_Erro=i
								url = nu_matricula&"_"&va_pt&"_"&erro
								grava = "no"
						end if					
					end if				
			else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "pt"
								matric_Erro=i
								url = nu_matricula&"_"&va_pt&"_"&erro
								grava = "no"
						end if
			end if
	end if


	if va_pp="" or isnull(va_pp) then
		va_pp=2		
		s_va_pp=0
		soma_va_pp=0		
	else
		teste_va_pp = isnumeric(va_pp)
		if teste_va_pp= true then					
		va_pp=va_pp*1			
					if va_pp =<100 then
						IF Int(va_pp)=va_pp THEN
							s_va_pp=1
							soma_va_pp=va_pp						
						ELSE	
							if  fail = 1 then
								grava = "no"
							else
								fail = 1 
								erro = "pp"
								matric_Erro=i
								url = nu_matricula&"_"&va_pp&"_"&erro
								grava = "no"
							end if					
						end if																				
					else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "pp"
								matric_Erro=i
								url = nu_matricula&"_"&va_pp&"_"&erro
								grava = "no"
						end if					
					end if				
			else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "pp"
								matric_Erro=i
								url = nu_matricula&"_"&va_pp&"_"&erro
								grava = "no"
						end if
			end if
	end if


'TESTES
s_va_t1=0
	if va_t1="" or isnull(va_t1) then
		va_t1=NULL		
		s_va_t1=0
		soma_va_t1=0		
	else
		teste_va_t1 = isnumeric(va_t1)
		if teste_va_t1= true then					
		va_t1=va_t1*1			
					if va_t1 =<100 then
						IF Int(va_t1)=va_t1 THEN
							s_va_t1=1
							soma_va_t1=va_t1						
						ELSE	
							if  fail = 1 then
								grava = "no"
							else
								fail = 1 
								erro = "t1"
								matric_Erro=i
								url = nu_matricula&"_"&va_av1&"_"&erro
								grava = "no"
							end if					
						end if																				
					else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "t1"
								matric_Erro=i
								url = nu_matricula&"_"&va_t1&"_"&erro
								grava = "no"
						end if					
					end if				
			else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "t1"
								matric_Erro=i
								url = nu_matricula&"_"&va_t1&"_"&erro
								grava = "no"
						end if
			end if
	end if
	if va_t2="" or isnull(va_t2) then
		va_t2=NULL		
		s_va_t2=0
		soma_va_t2=0		
	else
		teste_va_t2 = isnumeric(va_t2)
		if teste_va_t2= true then					
		va_t2=va_t2*1			
					if va_t2 =<100 then
						IF Int(va_t2)=va_t2 THEN
							s_va_t2=1
							soma_va_t2=va_t2						
						ELSE	
							if  fail = 1 then
								grava = "no"
							else
								fail = 1 
								erro = "t2"
								matric_Erro=i
								url = nu_matricula&"_"&va_t2&"_"&erro
								grava = "no"
							end if					
						end if																				
					else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "t2"
								matric_Erro=i
								url = nu_matricula&"_"&va_t2&"_"&erro
								grava = "no"
						end if					
					end if				
			else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "t2"
								matric_Erro=i
								url = nu_matricula&"_"&va_t2&"_"&erro
								grava = "no"
						end if
			end if
	end if
	if va_t3="" or isnull(va_t3) then
		va_t3=NULL		
		s_va_t3=0
		soma_va_t3=0		
	else
		teste_va_t3 = isnumeric(va_t3)
		if teste_va_t3= true then					
		va_t3=va_t3*1			
					if va_t3 =<100 then
						IF Int(va_t3)=va_t3 THEN
							s_va_t3=1
							soma_va_t3=va_t3						
						ELSE	
							if  fail = 1 then
								grava = "no"
							else
								fail = 1 
								erro = "va_t3"
								matric_Erro=i
								url = nu_matricula&"_"&va_t3&"_"&erro
								grava = "no"
							end if					
						end if																				
					else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "t3"
								matric_Erro=i
								url = nu_matricula&"_"&va_t3&"_"&erro
								grava = "no"
						end if					
					end if				
			else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "t3"
								matric_Erro=i
								url = nu_matricula&"_"&va_t3&"_"&erro
								grava = "no"
						end if
			end if
	end if

'//////////////////////////////////////////////////////////////////////
'Notas
s_va_p=0
	if va_p1="" or isnull(va_p1) then
		va_p1=NULL		
		s_va_p1=0
		soma_va_p1=0		
	else
		teste_va_p1 = isnumeric(va_p1)
		if teste_va_p1= true then					
		va_p1=va_p1*1			
					if va_p1 =<100 then
						IF Int(va_p1)=va_p1 THEN
							s_va_p1=1
							soma_va_p1=va_p1						
						ELSE	
							if  fail = 1 then
								grava = "no"
							else
								fail = 1 
								erro = "p1"
								matric_Erro=i
								url = nu_matricula&"_"&va_p1&"_"&erro
								grava = "no"
							end if					
						end if																				
					else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "p1"
								matric_Erro=i
								url = nu_matricula&"_"&va_p1&"_"&erro
								grava = "no"
						end if					
					end if				
			else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "p1"
								matric_Erro=i
								url = nu_matricula&"_"&va_p1&"_"&erro
								grava = "no"
						end if
			end if
	end if

	if va_p2="" or isnull(va_p2) then
		va_p2=NULL		
		s_va_p2=0
		soma_va_p2=0		
	else
		teste_va_p2 = isnumeric(va_p2)
		if teste_va_p2= true then					
		va_p2=va_p2*1			
					if va_p2 =<100 then
						IF Int(va_p2)=va_p2 THEN
							s_va_p2=1
							soma_va_p2=va_p2						
						ELSE	
							if  fail = 1 then
								grava = "no"
							else
								fail = 1 
								erro = "p2"
								matric_Erro=i
								url = nu_matricula&"_"&va_p2&"_"&erro
								grava = "no"
							end if					
						end if																				
					else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "p2"
								matric_Erro=i
								url = nu_matricula&"_"&va_p2&"_"&erro
								grava = "no"
						end if					
					end if				
			else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "p2"
								matric_Erro=i
								url = nu_matricula&"_"&va_p2&"_"&erro
								grava = "no"
						end if
			end if
	end if

	if va_p3="" or isnull(va_p3) then
		va_p3=NULL		
		s_va_p3=0
		soma_va_p3=0		
	else
		teste_va_p3 = isnumeric(va_p3)
		if teste_va_p3= true then					
		va_p3=va_p3*1			
					if va_p3 =<100 then
						IF Int(va_p3)=va_p3 THEN
							s_va_p3=1
							soma_va_p3=va_p3						
						ELSE	
							if  fail = 1 then
								grava = "no"
							else
								fail = 1 
								erro = "p3"
								matric_Erro=i
								url = nu_matricula&"_"&va_p3&"_"&erro
								grava = "no"
							end if					
						end if																				
					else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "p3"
								matric_Erro=i
								url = nu_matricula&"_"&va_p3&"_"&erro
								grava = "no"
						end if					
					end if				
			else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "p3"
								matric_Erro=i
								url = nu_matricula&"_"&va_p3&"_"&erro
								grava = "no"
						end if
			end if
	end if


	if va_bon="" or isnull(va_bon) then
		va_bon=NULL		
		s_va_bon=0
	else
		teste_va_bon = isnumeric(va_bon) 
		if teste_va_bon = true then
			if va_bon=<100 then
					va_bon=va_bon*1
						IF Int(va_bon)=va_bon THEN
							s_va_bon=va_bon													
						ELSE						
							if  fail = 1 then
								grava = "no"
							else												
								fail = 1 
								erro = "bon"
								url = nu_matricula&"_"&va_bon&"_"&erro
								grava = "no"
							end if					
						end if								
			else
						if  fail = 1 then
							grava = "no"
						else											
							fail = 1 
							erro = "bon"
							url = nu_matricula&"_"&va_bon&"_"&erro
							grava = "no"
						end if			
			end if

		else
						if  fail = 1 then
							grava = "no"
						else											
							fail = 1 
							erro = "bon"
							url = nu_matricula&"_"&va_bon&"_"&erro
							grava = "no"
						end if
		end if
	end if
	
	if va_rec="" or isnull(va_rec) then
		va_rec=NULL		
		s_va_rec=0		
	else
		teste_va_rec = isnumeric(va_rec) 
		if teste_va_rec = true then
			if va_rec=<100 then
						va_rec=va_rec*1
						IF Int(va_rec)=va_rec THEN
							s_va_rec=va_rec						
						ELSE	
							if  fail = 1 then
								grava = "no"
							else					
								fail = 1 
								erro = "rec"
								url = nu_matricula&"_"&va_rec&"_"&erro
								grava = "no"
							end if					
						end if								
			else
						if  fail = 1 then
							grava = "no"
						else					
							fail = 1 
							erro = "rec"
							url = nu_matricula&"_"&va_rec&"_"&erro
							grava = "no"
						end if							
			end if

		else
						if  fail = 1 then
							grava = "no"
						else					
							fail = 1 
							erro = "rec"
							url = nu_matricula&"_"&va_rec&"_"&erro
							grava = "no"
						end if
		end if
	end if	

'/////////////////////////////////////////////////////////////////////////
'Médias

	s_va_t1=s_va_t1*1
	s_va_t2=s_va_t2*1
	s_va_t3=s_va_t3*1
	soma_va_t1=soma_va_t1*1
	soma_va_t2=soma_va_t2*1
	soma_va_t3=soma_va_t3*1	
	
	divisor_mt=s_va_t1+s_va_t2+s_va_t3
	dividendo_mt=soma_va_t1+soma_va_t2+soma_va_t3
	
		if divisor_mt=0 or p_todas_mt_subs_lancadas = "N" THEN
			media_t="no"
			mt=NULL						
		else
			media_t="ok"
				mt=dividendo_mt/divisor_mt
				'm1=m1*10
					decimo = mt - Int(mt)
						If decimo >= 0.5 Then
							nota_arredondada = Int(mt) + 1
							mt=nota_arredondada
						'elseif decimo > 0 Then
						'	nota_arredondada = Int(m1) + 0.5
						'	m1=nota_arredondada
						else
							nota_arredondada = Int(mt)
							mt=nota_arredondada						
						End If
					'm1=m1/10				
					mt = formatNumber(mt,1)					
		end if

	s_va_p1=s_va_p1*1
	s_va_p2=s_va_p2*1
	s_va_p3=s_va_p3*1
	soma_va_p1=soma_va_p1*1
	soma_va_p2=soma_va_p2*1
	soma_va_p3=soma_va_p3*1	
	
	divisor_mp=s_va_p1+s_va_p2+s_va_p3
	dividendo_mp=soma_va_p1+soma_va_p2+soma_va_p3
	
		if divisor_mp=0 or p_todas_mp_subs_lancadas = "N" THEN
			media_p="no"
			mp=NULL						
		else
			media_p="ok"
				mp=dividendo_mp/divisor_mp
				'm1=m1*10
					decimo = mp - Int(mp)
						If decimo >= 0.5 Then
							nota_arredondada = Int(mp) + 1
							mp=nota_arredondada
						'elseif decimo > 0 Then
						'	nota_arredondada = Int(m1) + 0.5
						'	m1=nota_arredondada
						else
							nota_arredondada = Int(mp)
							mp=nota_arredondada						
						End If
					'm1=m1/10				
					mp = formatNumber(mp,1)					
		end if

	if media_t="ok" and media_p="ok" then
		m1=((mt*va_pt)+(mp*va_pp))/(va_pt+va_pp)
		'm1=m1*10
		decimo = m1 - Int(m1)
			If decimo >= 0.5 Then
				nota_arredondada = Int(m1) + 1
				m1=nota_arredondada
			'elseif decimo > 0 Then
			'	nota_arredondada = Int(m1) + 0.5
			'	m1=nota_arredondada
			else
				nota_arredondada = Int(m1)
				m1=nota_arredondada						
			End If
		'm1=m1/10							
		m1 = formatNumber(m1,1)
	else
		m1=NULL
	end if	

'response.Write(m1&"<BR>")

	if isnull(m1) or m1="" then
		m2=NULL
		m3=NULL	
	else		
		if isnull(va_bon) or va_bon="" then
		m2=m1		
		else
			m1=m1*1		
			va_bon=va_bon*1
			m2=m1+va_bon
			
			' if m2>100 then
				' if  fail = 1 then
					' grava = "no"
				' else											
					' fail = 1 
					' erro = "m2"
					' url = nu_matricula&"_"&va_bon&"_"&erro
					' grava = "no"
				' end if
			' end if
			'm2=m2*10
				decimo = m2 - Int(m2)
					If decimo >= 0.5 Then
						nota_arredondada = Int(m2) + 1
						m2=nota_arredondada
					'elseif decimo > 0 Then
					'	nota_arredondada = Int(m2) + 0.5
					'	m2=nota_arredondada
					else
						nota_arredondada = Int(m2)
						m2=nota_arredondada											
					End If
			'm2=m2/10				
				m2 = formatNumber(m2,1)
		end if
		
		if m2>100 then
			m2=100
		end if	
		
		m3=m2
		'if isnull(va_rec) or va_rec="" then
		'		decimo = m2 - Int(m2)
		'		If decimo > 0.5 Then
		'			nota_arredondada = Int(m2) + 1
		'			m2_arred=nota_arredondada
		'		elseIf decimo >= 0.25 Then
		'			nota_arredondada = Int(m2) + 0.5
		'			m2_arred=nota_arredondada
		'		else
		'			nota_arredondada = Int(m2)
		'			m2_arred=nota_arredondada											
		'		End If			
		'		m2_arred = formatNumber(m2_arred,1)
		'		m3=m2_arred						
		'else
		'	if periodo = 1 or periodo = 2 then					
		'			m2=m2*1
		'			va_rec=va_rec*1
		'			m3_temp=(m2+va_rec)/2
		'			decimo = m3_temp - Int(m3_temp)
		'				If decimo > 0.5 Then
		'					nota_arredondada = Int(m3_temp) + 1
		'					m3_temp=nota_arredondada
		'				elseIf decimo >= 0.25 Then
		'					nota_arredondada = Int(m3_temp) + 0.5
		'					m3_temp=nota_arredondada
		'				else
		'					nota_arredondada = Int(m3_temp)
		'					m3_temp=nota_arredondada											
		'				End If			
		'				m3_temp = formatNumber(m3_temp,1)
		'm2=m2*1
		'm3_temp=m3_temp*1						
					'if m3_temp >= m2 then
					'	m3=m3_temp
					'else
					'	decimo = m2 - Int(m2)
					'	If decimo > 0.5 Then
					'		nota_arredondada = Int(m2) + 1
					'		m2_arred=nota_arredondada
					'	elseIf decimo >= 0.25 Then
					'		nota_arredondada = Int(m2) + 0.5
					'		m2_arred=nota_arredondada
					'	else
					'		nota_arredondada = Int(m2)
					'		m2_arred=nota_arredondada											
					'	End If			
					'	m2_arred = formatNumber(m2_arred,1)
					'	m3=m2_arred
					'end if
			'else
			'		decimo = m2 - Int(m2)
			'		If decimo > 0.5 Then
			'			nota_arredondada = Int(m2) + 1
			'			m2_arred=nota_arredondada
			'		elseIf decimo >= 0.25 Then
			'			nota_arredondada = Int(m2) + 0.5
			'			m2_arred=nota_arredondada
			'		else
			'			nota_arredondada = Int(m2)
			'			m2_arred=nota_arredondada											
			'		End If			
			'		m2_arred = formatNumber(m2_arred,1)
			'		m3=m2_arred
			'end if							
		'end if
	end if

'	if nu_matricula = 20150114 then	
'response.Write(nu_matricula&"; "&co_materia_pr&"; "&co_materia&"; "&periodo&"; "&va_faltas&"; "&va_pt&"; "&va_pp&"; "&va_t1&"; "&va_t2&"; "&va_t3&"; "&va_p1&"; "&va_p2&"; "&va_p3&"; "&va_bon&"; "&va_rec&"; "&data&"; "&horario&"; "&co_usr&"<-<br>")
'response.Write(grava&";"&fail&";"&url&"<-<br>")
'response.End()
'end if
if grava = "ok" then
	
'	response.Write(va_pt&"<-<br>")
'response.Write(va_pp&"<-<br>")
'		response.Write("Select * from TB_Nota_A WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo)


'


		Set RS0 = Server.CreateObject("ADODB.Recordset")
		CONEXAO0 = "Select * from TB_Nota_A WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
		Set RS0 = CON.Execute(CONEXAO0)
		
				'response.Write(nu_matricula&" wrk_va_t1 "&va_t1&"<BR>")			
				'response.Write(nu_matricula&"wrk_va_t2 "&va_t2&"<BR>")	
				'response.Write("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++<BR>")		
					
	If RS0.EOF THEN	
	
		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Nota_A", CON, 2, 2 'which table do you want open
		RS.addnew
		
			RS("CO_Matricula") = nu_matricula
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Periodo") = periodo
			RS("NU_Faltas")=va_faltas				
			RS("VA_Teste1")=va_t1
			RS("VA_Teste2")=va_t2
			RS("VA_Teste3")=va_t3
			RS("MD_Teste")=mt
			RS("PE_Teste")=va_pt
			RS("VA_Prova1")=va_p1
			RS("VA_Prova2")=va_p2
			RS("VA_Prova3")=va_p3
			RS("MD_Prova")=mp
			RS("PE_Prova")=va_pp						
			RS("VA_Media1")=m1
			RS("VA_Bonus")=va_bon	
			RS("VA_Media2")=m2
			RS("VA_Rec")=va_rec
			RS("VA_Media3")=m3
			RS("DA_Ult_Acesso") = data
			RS("HO_ult_Acesso") = horario
			RS("CO_Usuario")= co_usr
		
		RS.update
		set RS=nothing
		response.Write(nu_matricula&" horario "&horario&"<BR>")
	else

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		CONEXAO0 = "DELETE * from TB_Nota_A WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
		Set RS0 = CON.Execute(CONEXAO0)

		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Nota_A", CON, 2, 2 'which table do you want open
		RS.addnew
		
			RS("CO_Matricula") = nu_matricula
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Periodo") = periodo
			RS("NU_Faltas")=va_faltas				
			RS("VA_Teste1")=va_t1
			RS("VA_Teste2")=va_t2
			RS("VA_Teste3")=va_t3
			RS("MD_Teste")=mt
			RS("PE_Teste")=va_pt
			RS("VA_Prova1")=va_p1
			RS("VA_Prova2")=va_p2
			RS("VA_Prova3")=va_p3
			RS("MD_Prova")=mp
			RS("PE_Prova")=va_pp						
			RS("VA_Media1")=m1
			RS("VA_Bonus")=va_bon	
			RS("VA_Media2")=m2
			RS("VA_Rec")=va_rec
			RS("VA_Media3")=m3
			RS("DA_Ult_Acesso") = data
			RS("HO_ult_Acesso") = horario
			RS("CO_Usuario")= co_usr

		
		RS.update
		set RS=nothing		
'		response.Write(nu_matricula&" co_materia "&co_materia&" co_materia_pr "&co_materia_pr&" va_p1 "&va_p1&"<BR>")
	end if
end if
'response.Write(fail)
'response.End()	
	if fail = 1 then
		Grava_BDA = co_materia_pr&"$!$"&url 
	else
		Grava_BDA = "S"		
	end if	
	
End Function
%>