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
		ABRIR = "DBQ="& CAMINHO_nd & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CONMT = Server.CreateObject("ADODB.Connection") 
		ABRIRMT = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONMT.Open ABRIRMT
		
				
		Set RSMT  = Server.CreateObject("ADODB.Recordset")
		SQL_MT  = "Select CO_Materia_Principal from TB_Materia WHERE CO_Materia = '"& co_materia&"'"
		Set RSMT  = CONMT.Execute(SQL_MT)

'response.Write(SQL_MT)
		
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
		response.Write(i&"-"& max&"<<BR>")	
	
	nu_matricula = request.form("nu_matricula_"&i)
'	response.Write(nu_matricula&"<BR>")
	if nu_matricula = "falta" then

			i=i*1
			max=max*1
		if i=max then
			grava = "no"
		else

			while nu_matricula = "falta"
				i=i+1
				nu_matricula = request.form("nu_matricula_"&i)
'				response.Write("i="&i&"<BR>")
'				response.Write(nu_matricula&"<BR>")							
			wend
				va_t1=request.form("t1_"&i)
				va_t2=request.form("t2_"&i)
				va_p1=request.form("p1_"&i)
				va_simul=request.form("simul_"&i)
				va_p2=request.form("p2_"&i)
				va_bon=request.form("bon_"&i)
				va_rec=request.form("rec_"&i)
		end if	
	else
		nu_matricula = request.form("nu_matricula_"&i)
		va_t1=request.form("t1_"&i)
		va_t2=request.form("t2_"&i)
		va_p1=request.form("p1_"&i)
		va_simul=request.form("simul_"&i)
		va_p2=request.form("p2_"&i)
		va_bon=request.form("bon_"&i)
		va_rec=request.form("rec_"&i)
	end if

	if fail = 0 then 
		Session("t1")=va_t1
		Session("t2")=va_t2
		Session("p1")=va_p1
		Session("simul")=va_simul
		Session("p2")=va_p2
		Session("bon")=va_bon
		Session("rec")=va_rec
	end if		
'////////////////////////////////////////////////////////////////
'pesos (por enquanto essa verifica��o n�o � usada)

	if va_pt="" or isnull(va_pt) then
		va_pt = 1
		p_va_pt="vazio"
		teste_va_pt= true
	else
		teste_va_pt = isnumeric(va_pt)
	end if

	if va_pp="" or isnull(va_pp) then
		va_pp = 2
		p_va_pp="vazio"
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



'//////////////////////////////////////////////////////////////////////
'Notas

	if va_simul="" or isnull(va_simul) then
		va_simul=NULL		
		s_va_simul=0
		soma_simul=0
		s_lancada="n"		
	else
		teste_va_simul = isnumeric(va_simul)
		if teste_va_simul= true then					
		va_simul=va_simul*1			
					if va_simul =<20 then
						IF Int(va_simul)=va_simul THEN
							s_va_simul=1
							soma_simul=va_simul
							s_lancada="ok"						
						ELSE
						s_lancada="n"	
							if  fail = 1 then
								grava = "no"
							else					
								fail = 1 
								erro = "simul"
								url = nu_matricula&"_"&va_simul&"_"&erro
								grava = "no"
							end if					
						end if					
					else
					s_lancada="n"					
						if  fail = 1 then
							grava = "no"
						else					
							fail = 1 
							erro = "simul"
							url = nu_matricula&"_"&va_simul&"_"&erro
							grava = "no"
						end if
					end if				
			else
			s_lancada="n"
						if  fail = 1 then
							grava = "no"
						else					
							fail = 1 
							erro = "simul"
							url = nu_matricula&"_"&va_simul&"_"&erro
							grava = "no"
						end if
			end if
	end if

s_va_p=0
	if s_lancada="ok" then
		limite_p1=80
		tipo_erro="p1b"
	else
		limite_p1=100	
		tipo_erro="p1"
	end if
		
	if va_p1="" or isnull(va_p1) then
		va_p1=NULL		
		s_va_p1=0
		soma_prova1=0
		p1_lancada="n"		
	else
		teste_va_p1 = isnumeric(va_p1)
		if teste_va_p1= true then					
			va_p1=va_p1*1	
			limite_p1=limite_p1*1					
					if va_p1 =<limite_p1 then
						IF Int(va_p1)=va_p1 THEN
							s_va_p1=1
							soma_prova1=va_p1
							p1_lancada="ok"						
						ELSE
						p1_lancada="n"	
							if  fail = 1 then
								grava = "no"
							else					
								fail = 1 
								erro = tipo_erro
								url = nu_matricula&"_"&va_p1&"_"&erro
								grava = "no"
							end if					
						end if																				
					else
					p1_lancada="n"
						if  fail = 1 then
							grava = "no"
						else					
							fail = 1 
							erro = tipo_erro
							url = nu_matricula&"_"&va_p1&"_"&erro
							grava = "no"
						end if					
					end if				
			else
				p1_lancada="n"
						if  fail = 1 then
							grava = "no"
						else					
							fail = 1 
							erro = tipo_erro
							url = nu_matricula&"_"&va_p1&"_"&erro
							grava = "no"
						end if					
			end if
	end if

	if va_p2="" or isnull(va_p2) then
		va_p2=NULL		
		s_va_p2=0
		soma_prova2=0
		p2_lancada="n"		
	else
		teste_va_p2 = isnumeric(va_p2)
		if teste_va_p2= true then					
		va_p2=va_p2*1			
					if va_p2 =<100 then					
						IF Int(va_p2)=va_p2 THEN
							s_va_p2=1
							soma_prova2=va_p2
							p2_lancada="ok"						
						ELSE
						p2_lancada="n"	
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
					p2_lancada="n"
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
			p2_lancada="n"
					if  fail = 1 then
							grava = "no"
						else
						fail = 1 
						erro = "p2"
						url = nu_matricula&"_"&va_p2&"_"&erro
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
	
	'response.Write(i&"-"&nu_matricula&"-"&va_t1 &"-"& va_t2&"-"&va_p1&"-"&va_simul&"-"&va_p2&"-"&va_bon&"-"&va_rec&"-<BR>")
	'response.Write(va_simul &" - "&fail&" - "&grava&" - "&erro&" <BR> ")	

'/////////////////////////////////////////////////////////////////////////
'M�dias


if grava = "ok" then
	soma_teste1=soma_teste1*1
	soma_teste2=soma_teste2*1
	s_va_t1=s_va_t1*1
	s_va_t2=s_va_t2*1
	
	s_va_t=s_va_t1+s_va_t2

		if teste_1_lancado="no" and teste_2_lancado="no" THEN
		media_t="no"
		mt=NULL
			
		else
		media_t="ok"		
		mt=(soma_teste1+soma_teste2)/s_va_t
		
		'mt=mt*10
			decimo = mt - Int(mt)
				If decimo >= 0.5 Then
					nota_arredondada = Int(mt) + 1
					mt=nota_arredondada
				Else
					nota_arredondada = Int(mt)
					mt=nota_arredondada					
				End If
			'mt=mt/10			
			mt = formatNumber(mt,0)					
		end if
		
		
	
	soma_prova1=soma_prova1*1
	soma_simul=soma_simul*1	
	soma_prova2=soma_prova2*1
	s_va_p1=s_va_p1*1
	s_va_simul=s_va_simul*1	
	s_va_p2=s_va_p2*1
	
	s_va_p=s_va_p1+s_va_simul

	'if s_va_p<2 and periodo<5 THEN
	'if media_t="ok" and ((p1_lancada="ok" and s_lancada="ok") or p2_lancada="ok") then
	if 	media_t="ok" and (p1_lancada="ok" or p2_lancada="ok") then		
		media_p="ok"		
			if periodo<5 then
				if p2_lancada="n" then		
					mp=soma_prova1+soma_simul
				elseif p1_lancada="n" and s_lancada="n" then
					mp=soma_prova2
				else
					mp=((soma_prova1+soma_simul)+soma_prova2)/2
				end if
				arredonda="sim"
			else			
				if isnull(soma_prova2) OR p2_lancada="n" then
				mp=NULL
				arredonda="nao"
				else
				mp=soma_prova2
				arredonda="sim"
				end if
			end if
		'mp=mp*10
			if arredonda="sim" then
				decimo = mp - Int(mp)
					If decimo >= 0.5 Then
						nota_arredondada = Int(mp) + 1
						mp=nota_arredondada
					Else
						nota_arredondada = Int(mp)
						mp=nota_arredondada					
					End If
				'mp=mp/10			
				mp = formatNumber(mp,0)
			end if				
	else			
		media_p="no"
		mp=NULL		
	end if

	if(co_materia_pr="PREA" and co_materia="PREA") or (co_materia_pr="PRE" and co_materia="PRE") or (co_materia_pr="PRS" and co_materia="PRS") or (co_materia_pr="PREC" and co_materia="PREC") or (co_materia_pr="DGEO" and co_materia="DGEO") or (co_materia_pr="ART1" and co_materia="ART1") or (co_materia_pr="EFIS" and co_materia="EFIS") or (co_materia_pr="EF" and co_materia="EF") or periodo>4 then
		if p1_lancada="n" then
			m1=NULL	
		else
			m1 = soma_prova1
			m1 = formatNumber(m1,0)		
		end if	
	else
		if 	media_t="ok" and media_p="ok" then
		va_pt=va_pt*1
		va_pp=va_pp*1
			if co_materia_pr="LP" and co_materia="LP" then
	
				if media_t="ok" and p1_lancada="ok" and s_lancada="ok" and p2_lancada="ok" then
				soma_prova1=soma_prova1*1
				soma_simul=soma_simul*1
				soma_prova2=soma_prova2*1
	'if i=2 then			
	'response.Write(mt&"-"&soma_prova1&"-"&soma_simul&"-"&soma_prova2)
	'response.end()				
	'end if
					m1=(mt+((soma_prova1+soma_simul)*2)+(soma_prova2*2))/5
						decimo = m1 - Int(m1)
							If decimo >= 0.5 Then
								nota_arredondada = Int(m1) + 1
								m1=nota_arredondada
							Else
								nota_arredondada = Int(m1)
								m1=nota_arredondada					
							End If
						'm1=m1/10			
						m1 = formatNumber(m1,0)
						
				ELSEif media_t="ok" and media_p="ok" and periodo =6 then
					mt=mt*1
					mp=mp*1
					m1=(mt+mp)/2
						decimo = m1 - Int(m1)
							If decimo >= 0.5 Then
								nota_arredondada = Int(m1) + 1
								m1=nota_arredondada
							Else
								nota_arredondada = Int(m1)
								m1=nota_arredondada					
							End If
						'm1=m1/10			
						m1 = formatNumber(m1,0)			
				
				ELSE		
								m1=NULL
				END IF
			ELSE						
				m1=((mt*va_pt)+(mp*va_pp))/(va_pt+va_pp)
				'm1=m1*10
					decimo = m1 - Int(m1)
						If decimo >= 0.5 Then
							nota_arredondada = Int(m1) + 1
							m1=nota_arredondada
						Else
							nota_arredondada = Int(m1)
							m1=nota_arredondada					
						End If
					'm1=m1/10			
					m1 = formatNumber(m1,0)		
			END IF
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
			'm2=m2*10
				decimo = m2 - Int(m2)
					If decimo >= 0.5 Then
						nota_arredondada = Int(m2) + 1
						m2=nota_arredondada
					Else
						nota_arredondada = Int(m2)
						m2=nota_arredondada					
					End If
			'm2=m2/10				
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
			curso=curso*1	
			etapa=etapa*1
'			if (curso=2 and etapa=1) or (curso=2 and etapa=2) then
'			Alterado por conta de e-mail de 11/05/2012
			if (curso=1 and etapa>1) or (curso=2) then			
				m2=m2*1
				va_rec=va_rec*1
				m3_temp=(m2+va_rec)/2
			else
				m2_temp=m2*2
				m2=m2*1
				va_rec=va_rec*1
				m3_temp=(m2_temp+va_rec)/3
				'response.Write(m3_temp &">"& m2)
				'response.End()
			end if	
			if m3_temp > m2 then
				m3=m3_temp				
				if m3>70 then
					m3=70
				end if
			else
				m3=m2
			end if
		

			
			'm3=m3*10
			decimo = m3 - Int(m3)
				If decimo >= 0.5 Then
					nota_arredondada = Int(m3) + 1
					m3=nota_arredondada
				Else
					nota_arredondada = Int(m3)
					m3=nota_arredondada					
				End If
			'm3=m3/10			
			m3 = formatNumber(m3,0)		
		end if
	end if
	
	
	'response.Write(nu_matricula&" - "&va_t1&" - "&va_pt&" - "&va_pp&"<BR>")
	
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		CONEXAO0 = "Select * from TB_Nota_D WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
		Set RS0 = CON.Execute(CONEXAO0)

	'response.Write(CONEXAO0&"<BR>")
		
	If RS0.EOF THEN	
	

		'response.Write("4"&turma &"/"&co_materia_pr)
		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Nota_D", CON, 2, 2 'which table do you want open
		RS.addnew
		
			RS("CO_Matricula") = nu_matricula
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Periodo") = periodo	
			RS("VA_Teste1")=va_t1
			RS("VA_Teste2")=va_t2
			RS("MD_Teste")=mt
			RS("PE_Teste")=va_pt
			RS("VA_Prova1")=va_p1
			RS("VA_Simul")=va_simul	
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
		CONEXAO0 = "DELETE * from TB_Nota_D WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
		Set RS0 = CON.Execute(CONEXAO0)
				'response.Write(CONEXAO0&"<BR>")

		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Nota_D", CON, 2, 2 'which table do you want open
		RS.addnew
		
			RS("CO_Matricula") = nu_matricula
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Periodo") = periodo	
			RS("VA_Teste1")=va_t1
			RS("VA_Teste2")=va_t2
			RS("MD_Teste")=mt
			RS("PE_Teste")=va_pt
			RS("VA_Prova1")=va_p1
			RS("VA_Simul")=va_simul	
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
		
	end if
end if
'response.Write(i&"_"&nu_matricula&"_"&fail&"<BR>")
next
if fail = 1 then
'response.Write(url)
response.Redirect(endereco_origem&"notas.asp?or=01&opt=err6&hp=err_"&url&"&obr="&obr) 

END IF

outro="P:"&periodo&",D:"&co_materia&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma&""

			call GravaLog (chave,outro)

if opt="cln" then
outro="P:"&periodo&",D:"&co_materia&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma&", Comunicou"
			call GravaLog (chave,outro)
response.Redirect(endereco_origem&"comunicar.asp?or=01&nota=TB_Nota_D&opt=ok&obr="&obr)
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