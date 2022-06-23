<%'On Error Resume Next%>
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
		ABRIR = "DBQ="& CAMINHO_nv & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON_TBF = Server.CreateObject("ADODB.Connection") 
		ABRIR_TBF = "DBQ="& CAMINHO_nf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_TBF.Open ABRIR_TBF
		
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

'	va_pt=request.form("pt")
'	va_pp1=request.form("pp1")
'	va_pp2=request.form("pp2")

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
			va_p1=request.form("p1_"&i)
			va_p2=request.form("p2_"&i)
			va_p3=request.form("p3_"&i)
			va_p4=request.form("p4_"&i)
			va_p5=request.form("p5_"&i)
			va_p6=request.form("p6_"&i)
			va_p7=request.form("p7_"&i)
			va_p8=request.form("p8_"&i)	
			va_p9=request.form("p9_"&i)
			va_p10=request.form("p10_"&i)								
			va_bon=request.form("bon_"&i)
			va_rec=request.form("rec_"&i)
		end if	
	else
		nu_matricula = request.form("nu_matricula_"&i)
		va_p1=request.form("p1_"&i)
		va_p2=request.form("p2_"&i)
		va_p3=request.form("p3_"&i)
		va_p4=request.form("p4_"&i)
		va_p5=request.form("p5_"&i)
		va_p6=request.form("p6_"&i)
		va_p7=request.form("p7_"&i)
		va_p8=request.form("p8_"&i)
		va_p9=request.form("p9_"&i)
		va_p10=request.form("p10_"&i)
		va_bon=request.form("bon_"&i)
		va_rec=request.form("rec_"&i)
	end if

	if fail = 0 then 
		Session("p1")=va_p1
		Session("p2")=va_p2
		Session("p3")=va_p3
		Session("p4")=va_p4
		Session("p5")=va_p5
		Session("p6")=va_p6
		Session("p7")=va_p7
		Session("p8")=va_p8
		Session("p9")=va_p9
		Session("p10")=va_p10
		Session("bon")=va_bon
		Session("rec")=va_rec
	end if		
'////////////////////////////////////////////////////////////////
'pesos 

'	if va_pt="" or isnull(va_pt) then
		va_pt = 1
'		p_va_pt="vazio"
'		teste_va_pt= true
'	else
'		teste_va_pt = isnumeric(va_pt)
'	end if
'
'	if va_pp1="" or isnull(va_pp1) then
		va_pp1 = 4
'		p_va_pp1="vazio"
'		teste_va_pp1= true
'	else
'		teste_va_pp1 = isnumeric(va_pp1)
'	end if
'
'
'	if va_pp2="" or isnull(va_pp2) then
		if periodo < 4 then 
			va_pp2 = 5
		else
			va_pp2 = 0		
		end if
'		p_va_pp2="vazio"
'		teste_va_pp2= true
'	else
'		teste_va_pp2 = isnumeric(va_pp1)
'	end if
'
'
'if teste_va_pt=true and teste_va_pp1=true and teste_va_pp2=true then
'va_pt=va_pt*1
'va_pp1=va_pp1*1
'va_pp2=va_pp2*1
''sum_p = va_pt+va_pp
''	if sum_p>100 then
''			fail = 1 
''			erro = "sp"
''			url = 0&"_"&sum_p&"_"&erro
''			grava = "no"
''	end if
'else
'			fail = 1 
'			erro = "pt"
'			url = 0&"_"&sum_p&"_"&erro
'			grava = "no"
'end if
'
'//////////////////////////////////////////////////////////////////////
'Notas

	limite_p1=100	
	tipo_erro="p1"

		
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
	
	if va_p3="" or isnull(va_p3) then
		va_p3=NULL		
		s_va_p3=0
		soma_prova3=0
		p3_lancada="n"		
	else
		teste_va_p3 = isnumeric(va_p3)
		if teste_va_p3= true then					
		va_p3=va_p3*1			
					if va_p3 =<100 then					
						IF Int(va_p3)=va_p3 THEN
							s_va_p3=1
							soma_prova3=va_p3
							p3_lancada="ok"						
						ELSE
						p3_lancada="n"	
							if  fail = 1 then
								grava = "no"
							else					
								fail = 1 
								erro = "p3"
								url = nu_matricula&"_"&va_p3&"_"&erro
								grava = "no"
							end if					
						end if					
					else
					p3_lancada="n"
						if  fail = 1 then
							grava = "no"
						else
							fail = 1 
							erro = "p3"
							url = nu_matricula&"_"&va_p3&"_"&erro
							grava = "no"
						end if
					end if				
			else
			p3_lancada="n"
					if  fail = 1 then
							grava = "no"
						else
						fail = 1 
						erro = "p3"
						url = nu_matricula&"_"&va_p3&"_"&erro
						grava = "no"
					end if
			end if
	end if
	
	if va_p4="" or isnull(va_p4) then
		va_p4=NULL		
		s_va_p4=0
		soma_prova4=0
		p4_lancada="n"		
	else
		teste_va_p4 = isnumeric(va_p4)
		if teste_va_p4= true then					
		va_p4=va_p4*1			
					if va_p4 =<100 then					
						IF Int(va_p4)=va_p4 THEN
							s_va_p4=1
							soma_prova4=va_p4
							p4_lancada="ok"						
						ELSE
						p4_lancada="n"	
							if  fail = 1 then
								grava = "no"
							else					
								fail = 1 
								erro = "p4"
								url = nu_matricula&"_"&va_p4&"_"&erro
								grava = "no"
							end if					
						end if					
					else
					p4_lancada="n"
						if  fail = 1 then
							grava = "no"
						else
							fail = 1 
							erro = "p4"
							url = nu_matricula&"_"&va_p4&"_"&erro
							grava = "no"
						end if
					end if				
			else
			p4_lancada="n"
					if  fail = 1 then
							grava = "no"
						else
						fail = 1 
						erro = "p4"
						url = nu_matricula&"_"&va_p4&"_"&erro
						grava = "no"
					end if
			end if
	end if	

	if va_p5="" or isnull(va_p5) then
		va_p5=NULL		
		s_va_p5=0
		soma_prova5=0
		p5_lancada="n"		
	else
		teste_va_p5 = isnumeric(va_p5)
		if teste_va_p5= true then					
		va_p5=va_p5*1			
					if va_p5 =<100 then					
						IF Int(va_p5)=va_p5 THEN
							s_va_p5=1
							soma_prova5=va_p5
							p5_lancada="ok"						
						ELSE
						p5_lancada="n"	
							if  fail = 1 then
								grava = "no"
							else					
								fail = 1 
								erro = "p5"
								url = nu_matricula&"_"&va_p5&"_"&erro
								grava = "no"
							end if					
						end if					
					else
					p5_lancada="n"
						if  fail = 1 then
							grava = "no"
						else
							fail = 1 
							erro = "p5"
							url = nu_matricula&"_"&va_p5&"_"&erro
							grava = "no"
						end if
					end if				
			else
			p5_lancada="n"
					if  fail = 1 then
							grava = "no"
						else
						fail = 1 
						erro = "p5"
						url = nu_matricula&"_"&va_p5&"_"&erro
						grava = "no"
					end if
			end if
	end if	
	if va_p6="" or isnull(va_p6) then
		va_p6=NULL		
		s_va_p6=0
		soma_prova6=0
		p6_lancada="n"		
	else
		teste_va_p6 = isnumeric(va_p6)
		if teste_va_p6= true then					
		va_p6=va_p6*1			
					if va_p6 =<100 then					
						IF Int(va_p6)=va_p6 THEN
							s_va_p6=1
							soma_prova6=va_p6
							p6_lancada="ok"						
						ELSE
						p6_lancada="n"	
							if  fail = 1 then
								grava = "no"
							else					
								fail = 1 
								erro = "p6"
								url = nu_matricula&"_"&va_p6&"_"&erro
								grava = "no"
							end if					
						end if					
					else
					p6_lancada="n"
						if  fail = 1 then
							grava = "no"
						else
							fail = 1 
							erro = "p6"
							url = nu_matricula&"_"&va_p6&"_"&erro
							grava = "no"
						end if
					end if				
			else
			p6_lancada="n"
					if  fail = 1 then
							grava = "no"
						else
						fail = 1 
						erro = "p6"
						url = nu_matricula&"_"&va_p6&"_"&erro
						grava = "no"
					end if
			end if
	end if		
	
	if va_p7="" or isnull(va_p7) then
		va_p7=NULL		
		s_va_p7=0
		soma_prova7=0
		p7_lancada="n"		
	else
		teste_va_p7 = isnumeric(va_p7)
		if teste_va_p7= true then					
		va_p7=va_p7*1			
					if va_p7 =<100 then					
						IF Int(va_p7)=va_p7 THEN
							s_va_p7=1
							soma_prova7=va_p7
							p7_lancada="ok"						
						ELSE
						p7_lancada="n"	
							if  fail = 1 then
								grava = "no"
							else					
								fail = 1 
								erro = "p7"
								url = nu_matricula&"_"&va_p7&"_"&erro
								grava = "no"
							end if					
						end if					
					else
					p7_lancada="n"
						if  fail = 1 then
							grava = "no"
						else
							fail = 1 
							erro = "p7"
							url = nu_matricula&"_"&va_p7&"_"&erro
							grava = "no"
						end if
					end if				
			else
			p7_lancada="n"
					if  fail = 1 then
							grava = "no"
						else
						fail = 1 
						erro = "p7"
						url = nu_matricula&"_"&va_p7&"_"&erro
						grava = "no"
					end if
			end if
	end if
	
	if va_p8="" or isnull(va_p8) then
		va_p8=NULL		
		s_va_p8=0
		soma_prova8=0
		p8_lancada="n"		
	else
		teste_va_p8 = isnumeric(va_p8)
		if teste_va_p8= true then					
		va_p8=va_p8*1			
					if va_p8 =<100 then					
						IF Int(va_p8)=va_p8 THEN
							s_va_p8=1
							soma_prova8=va_p8
							p8_lancada="ok"						
						ELSE
						p8_lancada="n"	
							if  fail = 1 then
								grava = "no"
							else					
								fail = 1 
								erro = "p8"
								url = nu_matricula&"_"&va_p8&"_"&erro
								grava = "no"
							end if					
						end if					
					else
					p8_lancada="n"
						if  fail = 1 then
							grava = "no"
						else
							fail = 1 
							erro = "p8"
							url = nu_matricula&"_"&va_p8&"_"&erro
							grava = "no"
						end if
					end if				
			else
			p8_lancada="n"
					if  fail = 1 then
							grava = "no"
						else
						fail = 1 
						erro = "p8"
						url = nu_matricula&"_"&va_p8&"_"&erro
						grava = "no"
					end if
			end if
	end if	

	if va_p9="" or isnull(va_p9) then
		va_p9=NULL		
		s_va_p9=0
		soma_prova9=0
		p9_lancada="n"		
	else
		teste_va_p9 = isnumeric(va_p9)
		if teste_va_p9= true then					
		va_p9=va_p9*1			
					if va_p9 =<100 then					
						IF Int(va_p9)=va_p9 THEN
							s_va_p9=1
							soma_prova9=va_p9
							p9_lancada="ok"						
						ELSE
						p9_lancada="n"	
							if  fail = 1 then
								grava = "no"
							else					
								fail = 1 
								erro = "p9"
								url = nu_matricula&"_"&va_p9&"_"&erro
								grava = "no"
							end if					
						end if					
					else
					p9_lancada="n"
						if  fail = 1 then
							grava = "no"
						else
							fail = 1 
							erro = "p9"
							url = nu_matricula&"_"&va_p9&"_"&erro
							grava = "no"
						end if
					end if				
			else
			p9_lancada="n"
					if  fail = 1 then
							grava = "no"
						else
						fail = 1 
						erro = "p9"
						url = nu_matricula&"_"&va_p9&"_"&erro
						grava = "no"
					end if
			end if
	end if
	
	if va_p10="" or isnull(va_p10) then
		va_p10=NULL		
		s_va_p10=0
		soma_prova10=0
		p10_lancada="n"		
	else
		teste_va_p10 = isnumeric(va_p10)
		if teste_va_p10= true then					
		va_p10=va_p10*1			
					if va_p10 =<100 then					
						IF Int(va_p10)=va_p10 THEN
							s_va_p10=1
							soma_prova10=va_p10
							p10_lancada="ok"						
						ELSE
						p10_lancada="n"	
							if  fail = 1 then
								grava = "no"
							else					
								fail = 1 
								erro = "p10"
								url = nu_matricula&"_"&va_p10&"_"&erro
								grava = "no"
							end if					
						end if					
					else
					p10_lancada="n"
						if  fail = 1 then
							grava = "no"
						else
							fail = 1 
							erro = "p10"
							url = nu_matricula&"_"&va_p10&"_"&erro
							grava = "no"
						end if
					end if				
			else
			p10_lancada="n"
					if  fail = 1 then
							grava = "no"
						else
						fail = 1 
						erro = "p10"
						url = nu_matricula&"_"&va_p10&"_"&erro
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
'Médias


if grava = "ok" then
soma_notas = s_va_p1+s_va_p2+s_va_p3+s_va_p4+s_va_p5+s_va_p6+s_va_p7+s_va_p8+s_va_p9+s_va_p10
	if soma_notas>=5 then


		
		
	
	soma_prova1=soma_prova1*1
	soma_prova2=soma_prova2*1
	soma_prova3=soma_prova3*1
	soma_prova4=soma_prova4*1
	soma_prova5=soma_prova5*1
	soma_prova6=soma_prova6*1
	soma_prova7=soma_prova7*1	
	soma_prova8=soma_prova8*1
	soma_prova9=soma_prova9*1
	soma_prova10=soma_prova10*1						
	s_va_p1=s_va_p1*1
	s_va_p2=s_va_p2*1
	s_va_p3=s_va_p3*1
	s_va_p4=s_va_p4*1	
	s_va_p5=s_va_p5*1	
	s_va_p6=s_va_p6*1	
	s_va_p7=s_va_p7*1	
	s_va_p8=s_va_p8*1	
	s_va_p9=s_va_p9*1	
	s_va_p10=s_va_p10*1	
							


			if periodo<>4 then
					m1=(soma_prova1+soma_prova2+soma_prova3+soma_prova4+soma_prova5+soma_prova6+soma_prova7+soma_prova8+soma_prova9+soma_prova10)/soma_notas
			
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
			end if				
	else			
		m1=NULL	
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
			
		if isnull(va_rec) or va_rec="" then
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
	
	if periodo<4 then
		va_prova2=m1
	else
		va_prova2=NULL	
	end if	
	
	'response.Write(nu_matricula&" - "&va_t1&" - "&va_pt&" - "&va_pp&"<BR>")
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		CONEXAO0 = "Select * from TB_Nota_V WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
		Set RS0 = CON.Execute(CONEXAO0)

	'response.Write(CONEXAO0&"<BR>")
		
	If RS0.EOF THEN	
	

		'response.Write("4"&turma &"/"&co_materia_pr)
		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Nota_V", CON, 2, 2 'which table do you want open
		RS.addnew
		
			RS("CO_Matricula") = nu_matricula
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Periodo") = periodo	
			RS("VA_PT1")=va_p1
			RS("VA_PT2")=va_p2
			RS("VA_PT3")=va_p3
			RS("VA_PT4")=va_p4
			RS("VA_PT5")=va_p5
			RS("VA_PT6")=va_p6
			RS("VA_PT7")=va_p7
			RS("VA_PT8")=va_p8
			RS("VA_PT9")=va_p9
			RS("VA_PT10")=va_p10
			RS("MD_PT")=m1
			RS("DA_Ult_Acesso") = data
			RS("HO_ult_Acesso") = horario
			RS("CO_Usuario")= co_usr
		
		RS.update
		set RS=nothing
		
	else
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		CONEXAO0 = "DELETE * from TB_Nota_V WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
		Set RS0 = CON.Execute(CONEXAO0)
				'response.Write(CONEXAO0&"<BR>")

		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Nota_V", CON, 2, 2 'which table do you want open
		RS.addnew
		
			RS("CO_Matricula") = nu_matricula
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Periodo") = periodo	
			RS("VA_PT1")=va_p1
			RS("VA_PT2")=va_p2
			RS("VA_PT3")=va_p3
			RS("VA_PT4")=va_p4
			RS("VA_PT5")=va_p5
			RS("VA_PT6")=va_p6
			RS("VA_PT7")=va_p7
			RS("VA_PT8")=va_p8
			RS("VA_PT9")=va_p9
			RS("VA_PT10")=va_p10
			RS("MD_PT")=m1
			RS("DA_Ult_Acesso") = data
			RS("HO_ult_Acesso") = horario
			RS("CO_Usuario")= co_usr
		
		RS.update
		set RS=nothing		
		
	end if	
	
	
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		CONEXAO0 = "Select * from TB_Nota_F WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
		Set RS0 = CON_TBF.Execute(CONEXAO0)

	'response.Write(CONEXAO0&"<BR>")
		
	If RS0.EOF THEN	
	

		'response.Write("4"&turma &"/"&co_materia_pr)
		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Nota_F", CON_TBF, 2, 2 'which table do you want open
		RS.addnew
		
			RS("CO_Matricula") = nu_matricula
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Periodo") = periodo	
			RS("VA_Teste1")=m1
			RS("VA_Teste2")=m1
			RS("MD_Teste")=m1
			RS("PE_Teste")=va_pt
			RS("VA_Prova1")=m1
			RS("VA_Prova2")=va_prova2
			RS("PE_Prova1")=va_pp1
			RS("PE_Prova2")=va_pp2
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
		CONEXAO0 = "DELETE * from TB_Nota_F WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
		Set RS0 = CON_TBF.Execute(CONEXAO0)
				'response.Write(CONEXAO0&"<BR>")

		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Nota_F", CON_TBF, 2, 2 'which table do you want open
		RS.addnew
		
			RS("CO_Matricula") = nu_matricula
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Periodo") = periodo	
			RS("VA_Teste1")=m1
			RS("VA_Teste2")=m1
			RS("MD_Teste")=m1
			RS("PE_Teste")=va_pt
			RS("VA_Prova1")=m1
			RS("VA_Prova2")=va_prova2
			RS("PE_Prova1")=va_pp1
			RS("PE_Prova2")=va_pp2
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
response.Redirect(endereco_origem&"comunicar.asp?or=01&nota=TB_Nota_F&opt=ok&obr="&obr)
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