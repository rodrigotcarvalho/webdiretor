<%'On Error Resume Next%>
<!--#include file="caminhos.asp"-->
<!--#include file="caminhos.asp"-->
<!--#include file="funcoes.asp"-->
<%
call cabecalho(1)
chave = session("chave")
session("chave")=chave
split_chave=split(chave,"-")
sistema_origem=split_chave(0)

if sistema_origem="WN" then
	endereco_origem="../wn/lancar/notas/lancar/"
elseif sistema_origem="WA" then	
	endereco_origem="../wa/professor/cna/notas/"
end if	

opt=request.QueryString("opt")
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 
data = dia &"/"& mes &"/"& ano
horario = hora & ":"& min

co_materia = request.form("co_materia")
unidade= request.form("unidade")
curso= request.form("curso")
etapa= request.form("etapa")
turma= request.form("turma")
periodo = request.form("periodo")
ano_letivo = request.form("ano_letivo")
co_prof = request.form("co_prof")
co_usr = request.form("co_usr")
max = request.form("max")

obr=co_materia&"?"&unidade&"?"&curso&"?"&etapa&"?"&turma&"?"&periodo&"?"&ano_letivo&"?"&co_prof

i=1

Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_nc & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CONMT = Server.CreateObject("ADODB.Connection") 
		ABRIRMT = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONMT.Open ABRIRMT
		
				
		Set RSMT  = Server.CreateObject("ADODB.Recordset")
		SQL_MT  = "Select CO_Materia_Principal from TB_Materia WHERE CO_Materia = '"& co_materia&"'"
		Set RSMT  = CONMT.Execute(SQL_MT)
		
co_materia_pr = RSMT("CO_Materia_Principal")
		
if Isnull(co_materia_pr) then
co_materia_pr= co_materia
else
co_materia_pr = co_materia_pr
end if

	va_pt=request.form("pt")
	va_pp=request.form("pp")

fail = 0
for i=1 to max
	grava="ok"
		response.Write(i&"<<BR>")	
	
	nu_matricula = request.form("nu_matricula_"&i)
	
		'response.Write(nu_matricula&"<<BR>")	
	if nu_matricula = "falta" then
			i=i*1
			max=max*1
		if i=max then
			grava = "no"
		else
			while nu_matricula = "falta"
				i=i+1
'Pega todas as vari�veis, pois quando nu_matricula for <> de falta o pr�ximo aluno j� ter� notas a serem buscadas.		
				nu_matricula = request.form("nu_matricula_"&i)
			wend
				va_t1=request.form("t1_"&i)
				va_t2=request.form("t2_"&i)
				va_t3=request.form("t3_"&i)
				va_t4=request.form("t4_"&i)
				va_p1=request.form("p1_"&i)
				va_p2=request.form("p2_"&i)
				va_bon=request.form("bon_"&i)
				va_rec=request.form("rec_"&i)			
		end if	
	else
		nu_matricula = request.form("nu_matricula_"&i)
		va_t1=request.form("t1_"&i)
		va_t2=request.form("t2_"&i)
		va_t3=request.form("t3_"&i)
		va_t4=request.form("t4_"&i)
		va_p1=request.form("p1_"&i)
		va_p2=request.form("p2_"&i)
		va_bon=request.form("bon_"&i)
		va_rec=request.form("rec_"&i)
	end if	
		
	if fail = 0 then 	
		Session("va_t1")=va_t1
		Session("va_t2")=va_t2
		Session("va_t3")=va_t3
		Session("va_t4")=va_t4
		Session("va_p1")=va_p1
		Session("va_p2")=va_p2
		Session("va_bon")=va_bon
		Session("va_rec")=va_rec	
	end if	
'////////////////////////////////////////////////////////////////
'pesos (por enquanto essa verifica��o n�o � usada)

	if va_pt="" or isnull(va_pt) then
		va_pt = 1
		'p_va_pt="vazio"
		teste_va_pt= true
	else
		teste_va_pt = isnumeric(va_pt)
	end if

	if va_pp="" or isnull(va_pp) then
		va_pp = 1
		'p_va_pp="vazio"
		teste_va_pp= true
	else
		teste_va_pp = isnumeric(va_pp)
	end if

	if teste_va_pt=true and teste_va_pp=true then
	va_pt=va_pt*1
	va_pp=va_pp*1
	
	'sum_p = va_pt+va_pp
	'	if sum_p>100 then
	'			fail = 1 
	'			erro = "sp"
	'			url = 0&"_"&sum_p&"_"&erro
	'			grava = "no"
	'	end if
	
	else
				fail = 1 
				erro = "pt"
				url = 0&"_"&sum_p&"_"&erro
				grava = "no"
	end if
	
	'///////////////////////////////////////////////////////////////////////////
	
	'TESTES
	s_va_t=0
		if va_t1="" or isnull(va_t1) then
			va_t1=NULL		
			s_va_t1=0
			soma_teste1=0	
			teste_1_lancado="no"	
		else
			teste_va_t1 = isnumeric(va_t1)
			if teste_va_t1= true then					
			va_t1=va_t1*1			
						if va_t1 =<100 then
							IF Int(va_t1)=va_t1 THEN
								s_va_t1=1
								soma_teste1=va_t1
								teste_1_lancado="sim"							
							ELSE	
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
				else
							if  fail = 1 then
								grava = "no"
							else
								fail = 1 
								erro = "t1"
								url = nu_matricula&"_"&va_t1&"_"&erro
								grava = "no"
							end if
				end if
		end if
	
		if va_t2="" or isnull(va_t2) then
			va_t2=NULL		
			s_va_t2=0
			soma_teste2=0	
			teste_2_lancado="no"		
		else
			teste_va_t2 = isnumeric(va_t2)
			if teste_va_t2= true then					
			va_t2=va_t2*1			
						if va_t2 =<100 then			
							IF Int(va_t2)=va_t2 THEN
								s_va_t2=1
								soma_teste2=va_t2	
								teste_2_lancado="sim"													
							ELSE	
								if  fail = 1 then
									grava = "no"
								else					
									fail = 1 
									erro = "t2"
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
								url = nu_matricula&"_"&va_t2&"_"&erro
								grava = "no"
							end if
				end if
		end if	
	
		
		
		if va_t3="" or isnull(va_t3) then
			va_t3=NULL		
			s_va_t3=0
			soma_teste3=0
			teste_3_lancado="no"					
		else
			teste_va_t3 = isnumeric(va_t3)
			if teste_va_t3= true then					
			va_t3=va_t3*1			
						if va_t3 =<100 then
							IF Int(va_t3)=va_t3 THEN
								s_va_t3=1
								soma_teste3=va_t3
								teste_3_lancado="sim"														
							ELSE	
								if  fail = 1 then
									grava = "no"
								else					
									fail = 1 
									erro = "t3"
									url = nu_matricula&"_"&va_t3&"_"&erro
									grava = "no"
								end if					
							end if																				
						else
						fail = 1 
						erro = "t3"
						url = nu_matricula&"_"&va_t3&"_"&erro
						grava = "no"
						end if				
				else
				fail = 1 
				erro = "t3"
				url = nu_matricula&"_"&va_t3&"_"&erro
				grava = "no"
				end if
		end if
		
		if va_t4="" or isnull(va_t4) then
			va_t4=NULL		
			s_va_t4=0
			soma_teste4=0
			teste_4_lancado="no"					
		else
			teste_va_t4 = isnumeric(va_t4)
			if teste_va_t4= true then					
			va_t4=va_t4*1			
						if va_t4 =<100 then
							IF Int(va_t4)=va_t4 THEN
								s_va_t4=1
								soma_teste4=va_t4
								teste_4_lancado="sim"														
							ELSE	
								if  fail = 1 then
									grava = "no"
								else					
									fail = 1 
									erro = "t4"
									url = nu_matricula&"_"&va_t4&"_"&erro
									grava = "no"
								end if					
							end if																				
						
																					
						else
						fail = 1 
						erro = "t4"
						url = nu_matricula&"_"&va_t4&"_"&erro
						grava = "no"
						end if				
				else
				fail = 1 
				erro = "t4"
				url = nu_matricula&"_"&va_t4&"_"&erro
				grava = "no"
				end if
		end if	
	
	'response.Write(i&"-"&nu_matricula&"-"&va_apr7 &">"& va_v_apr7&"<BR>")
	
	'//////////////////////////////////////////////////////////////////////
	'Notas
	s_va_p=0
		if va_p1="" or isnull(va_p1) then
			va_p1=NULL		
			s_va_p1=0
			soma_prova1=0	
			prova_1_lancado="no"			
		else
			teste_va_p1 = isnumeric(va_p1)
			if teste_va_p1= true then					
			va_p1=va_p1*1			
						if va_p1 =<100 then
							IF Int(va_p1)=va_p1 THEN
								s_va_p1=1
								soma_prova1=va_p1	
								prova_1_lancado="sim"												
							ELSE	
								if  fail = 1 then
									grava = "no"
								else					
									fail = 1 
									erro = "p1"
									url = nu_matricula&"_"&va_p1&"_"&erro
									grava = "no"
								end if					
							end if															
						else
						fail = 1 
						erro = "p1"
						url = nu_matricula&"_"&va_p1&"_"&erro
						grava = "no"
						end if				
				else
				fail = 1 
				erro = "p1"
				url = nu_matricula&"_"&va_p1&"_"&erro
				grava = "no"
				end if
		end if
	
		if va_p2="" or isnull(va_p2) then
			va_p2=NULL		
			s_va_p2=0
			soma_prova2=0		
			prova_2_lancado="no"		
		else
			teste_va_p2 = isnumeric(va_p2)
			if teste_va_p2= true then					
			va_p2=va_p2*1			
						if va_p2 =<100 then
							IF Int(va_p2)=va_p2 THEN
								s_va_p2=1
								soma_prova2=va_p2	
								prova_2_lancado="sim"												
							ELSE	
								if  fail = 1 then
									grava = "no"
								else					
									fail = 1 
									erro = "p2"
									url = nu_matricula&"_"&va_p2&"_"&erro
									grava = "no"
								end if					
							end if															
						else
						fail = 1 
						erro = "p2"
						url = nu_matricula&"_"&va_p2&"_"&erro
						grava = "no"
						end if				
				else
				fail = 1 
				erro = "p2"
				url = nu_matricula&"_"&va_p2&"_"&erro
				grava = "no"
				end if
		end if
	
		if va_bon="" or isnull(va_bon) then
			va_bon=NULL		
			s_va_bon=0
		else
			teste_va_bon = isnumeric(va_bon) 
			if teste_va_bon = true then
				if va_bon<=100 then
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
				if va_rec<=100 then
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
	'M�dias
	
	
	if grava = "ok" then
		soma_teste1=soma_teste1*1
		soma_teste2=soma_teste2*1
		soma_teste3=soma_teste3*1
		soma_teste4=soma_teste4*1
		s_va_t1=s_va_t1*1
		s_va_t2=s_va_t2*1
		s_va_t3=s_va_t3*1
		s_va_t4=s_va_t4*1
		
		soma_teste=soma_teste1+soma_teste2+soma_teste3+soma_teste4	
			if teste_1_lancado="no" and teste_2_lancado="no" and teste_3_lancado="no" and teste_4_lancado="no" THEN
			media_t="no"
			mt=NULL						
			else
				if soma_teste =<100 then
					if  fail = 1 then
						mt=NULL
						grava = "no"
					else
					mt=soma_teste					
						decimo = mt - Int(mt)
						If decimo >= 0.5 Then
							nota_arredondada = Int(mt) + 1
							mt=nota_arredondada
						Else
							nota_arredondada = Int(mt)
							mt=nota_arredondada					
						End If			
						mt = formatNumber(mt,0)
						media_t="ok"
					end if																								
				else
					media_t="no"					
					mt=NULL
					fail = 1 
					erro = "mt"
					url = nu_matricula&"_"&s_va_t&"_"&erro
					grava = "no"
				end if				
			end if
		
		'response.End()	
		
		soma_prova1=soma_prova1*1
		soma_prova2=soma_prova2*1
		soma_prova3=soma_prova3*1
		s_va_p1=s_va_p1*1
		s_va_p2=s_va_p2*1
		s_va_p3=s_va_p3*1
		
		s_va_p=s_va_p1+s_va_p2
		
		IF s_va_p=0 THEN
			s_va_p=1
		END IF		
		
	'response.Write("if "&prova_1_lancado&"=no and "&prova_2_lancado&"=no MP='"&mp&"'<BR>")
	
		if (prova_1_lancado="no" AND prova_2_lancado="no") THEN
		media_p="no"
		mp=NULL		
		else
			if (periodo>3 and session("ano_letivo")>=2017) then
				mp=soma_prova1
			else
				mp=(soma_prova1+soma_prova2+soma_prova3)/s_va_p
		'mp=mp*10
				decimo = mp - Int(mp)
				If decimo >= 0.5 Then
					nota_arredondada = Int(mp) + 1
					mp=nota_arredondada
				Else
					nota_arredondada = Int(mp)
					mp=nota_arredondada					
				End If
			End If					
		'	mp=mp/10			
			mp = formatNumber(mp,0)
			media_p="ok"				
		end if

		if (co_materia_pr="EART" and co_materia="EART") or (co_materia_pr="EFIS" and co_materia="EFIS") or (co_materia_pr="ARTC" and co_materia="ARTC") or (periodo>4 and session("ano_letivo")<2017) or (periodo>3 and session("ano_letivo")>=2017) then
			if prova_1_lancado="no" then
				m1=NULL	
			else
				m1 = soma_prova1
				m1 = formatNumber(m1,0)		
			end if	
		else
			if 	media_t="ok" and media_p="ok" then
			va_pt=va_pt*1
			va_pp=va_pp*1
				m1=((mt*va_pt)+(mp*va_pp))/(va_pt+va_pp)
				response.Write(" mt = "&mt&" pt = "&va_pt&"<BR>")
				
				'm1=m1*10
					decimo = m1 - Int(m1)
						If decimo >= 0.5 Then
							nota_arredondada = Int(m1) + 1
							m1=nota_arredondada
						Else
							nota_arredondada = Int(m1)
							m1=nota_arredondada					
						End If
				'	m1=m1/10			
					m1 = formatNumber(m1,0)		
			else
				m1=NULL		
			END IF
		end if	
		
	
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
				
				if m2>100 then
				m2=100
				end if
		'		m2=m2*10
					decimo = m2 - Int(m2)
						If decimo >= 0.5 Then
							nota_arredondada = Int(m2) + 1
							m2=nota_arredondada
						Else
							nota_arredondada = Int(m2)
							m2=nota_arredondada					
						End If
			'	m2=m2/10				
					m2 = formatNumber(m2,0)
			end if
	
	
		if va_rec="" or isnull(va_rec) then
			va_rec=NULL		
			s_va_rec=0		
		else
			if session("ano_letivo")>=2017 then
				if isnumeric(va_rec) and isnumeric(m2) then
					if m2>=70 then
						if  fail = 1 then
							grava = "no"
						else					
							fail = 1 
							erro = "rec70"
							url = nu_matricula&"_"&va_rec&"_"&erro
							grava = "no"
						end if					
					end if	
				end if											
			end if
		end if
		
		if s_va_rec=0 or (m2>70 and session("ano_letivo")>=2017) then
				m3=m2			
			else					
				m2=m2*1
				va_rec=va_rec*1
				m3_temp=(m2+va_rec)/2
				mt=mt*1
				mp=mp*1
				
	'Inclu�do por conta de e-mail de 11/05/2012
	'================================================
				if (curso=1 and etapa>1) or (curso=2) then						
					if m3_temp > m2 then
						m3=m3_temp				
						if m3>70 then
							m3=70
						end if
					else
						m3=m2
					end if	
				else				
	'FIM da altera��o========================================			
					'response.Write(mt&">"&mp &"and"& mp&"<"&va_rec)
					if mt<=mp and mt<va_rec then
					'response.Write("1<BR>")
						m3=(((va_rec*va_pt)+(mp*va_pp))/(va_pt+va_pp))+s_va_bon
					elseif mt>mp and mp<va_rec then
					'response.Write("<BR>2<BR>")
						m3=(((mt*va_pt)+(va_rec*va_pp))/(va_pt+va_pp))+s_va_bon
					'response.Write("'"&m3&"=((("&mt&"*"&va_pt&")+("&va_rec&"*"&va_pp&"))/("&va_pt&"+"&va_pp&"))+"&s_va_bon)
					else
					'response.Write("3<BR>")
						m3=(((mt*va_pt)+(mp*va_pp))/(va_pt+va_pp))+s_va_bon
					end if			
					'm3=m3*10
	'Inclu�do por conta de e-mail de 11/05/2012
	'================================================					
				end if
	'FIM da altera��o========================================				
				decimo = m3 - Int(m3)
					If decimo >= 0.5 Then
						nota_arredondada = Int(m3) + 1
						m3=nota_arredondada
					Else
						nota_arredondada = Int(m3)
						m3=nota_arredondada					
					End If
				'm3=m3/10
				'response.Write("'"&va_rec&"'")			
				m3 = formatNumber(m3,0)		
			end if
		end if
		
			Set RS0 = Server.CreateObject("ADODB.Recordset")
			CONEXAO0 = "Select * from TB_Nota_C WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
			response.Write(CONEXAO0&"<BR>")
			Set RS0 = CON.Execute(CONEXAO0)
			
	
			
		If RS0.EOF THEN	
		
			
			'response.Write("4"&turma &"/"&co_materia_pr)
			Set RS = server.createobject("adodb.recordset")
			
			RS.open "TB_Nota_C", CON, 2, 2 'which table do you want open
			RS.addnew
			
				RS("CO_Matricula") = nu_matricula
				RS("CO_Materia_Principal") = co_materia_pr
				RS("CO_Materia") = co_materia
				RS("NU_Periodo") = periodo	
				RS("VA_Teste1")=va_t1
				RS("VA_Teste2")=va_t2
				RS("VA_Teste3")=va_t3
				RS("VA_Teste4")=va_t4
				RS("MD_Teste")=mt
				RS("PE_Teste")=va_pt
				RS("VA_Prova1")=va_p1
				RS("VA_Prova2")=va_p2	
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
			
		else
			
			Set RS0 = Server.CreateObject("ADODB.Recordset")
			CONEXAO0 = "DELETE * from TB_Nota_C WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
			Set RS0 = CON.Execute(CONEXAO0)
	
			Set RS = server.createobject("adodb.recordset")
			
			RS.open "TB_Nota_C", CON, 2, 2 'which table do you want open
			RS.addnew
			
				RS("CO_Matricula") = nu_matricula
				RS("CO_Materia_Principal") = co_materia_pr
				RS("CO_Materia") = co_materia
				RS("NU_Periodo") = periodo	
				RS("VA_Teste1")=va_t1
				RS("VA_Teste2")=va_t2
				RS("VA_Teste3")=va_t3
				RS("VA_Teste4")=va_t4
				RS("MD_Teste")=mt
				RS("PE_Teste")=va_pt
				RS("VA_Prova1")=va_p1
				RS("VA_Prova2")=va_p2	
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
			
	'		sql_atualiza= "UPDATE TB_Nota_C SET VA_Teste1 ="&sql_va_t1&", VA_Teste2 ="&sql_va_t2&", VA_Teste3 ="&sql_va_t3&", VA_Teste4 ="&sql_va_t4&", MD_Teste =FORMAT("&sql_mt&",2), "
	'		sql_atualiza=sql_atualiza&"PE_Teste ="&va_pt&", VA_Prova1 ="&sql_va_p1&", VA_Prova2 ="&sql_va_p2&", VA_Prova3="&sql_va_p3&", MD_Prova ="&sql_mp&", PE_Prova ="&va_pp&", VA_Media1 ="&sql_m1&", "
	'		sql_atualiza=sql_atualiza&"VA_Bonus ="&sql_va_bon&", VA_Media2 ="&sql_m2&", VA_Rec ="&sql_va_rec&", VA_Media3 ="&sql_m3&", "
	'		sql_atualiza=sql_atualiza&" DA_Ult_Acesso =#"& data &"#, HO_ult_Acesso =#"& horario &"#, CO_Usuario="& co_usr &"  WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
	
			
	'		response.Write(sql_atualiza&"<<6")
	'response.end()
	'		Set RS2 = Con.Execute(sql_atualiza)
			
		end if
	end if
	'response.Write(i&"-grava-"&grava&"hp=err_"&url&"&obr="&obr&"<br>")
next
if fail = 1 then

response.Redirect(endereco_origem&"notas.asp?or=01&opt=err6&hp=err_"&url&"&obr="&obr) 

END IF



'response.Write(">>>"&periodo)




outro="P:"&periodo&",D:"&co_materia&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma&""

			call GravaLog (chave,outro)

if opt="cln" then
outro="P:"&periodo&",D:"&co_materia&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma&", Comunicou"
			call GravaLog (chave,outro)
response.Redirect(endereco_origem&"comunicar.asp?or=01&nota=TB_Nota_C&opt=ok&obr="&obr)
else
response.Redirect(endereco_origem&"notas.asp?or=01&opt=ok&obr="&obr)
end if

%>
<%If Err.number<>0 then
errnumb = Err.number
errdesc = Err.Description
lsPath = Request.ServerVariables("SCRIPT_NAME")
arPath = Split(lsPath, "/")
GetFileName =arPath(UBound(arPath,1))
passos = 0
for way=0 to UBound(arPath,1)
passos=passos+1
next
seleciona1=passos-2
pasta=arPath(seleciona1)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("erro.asp")
end if
%>