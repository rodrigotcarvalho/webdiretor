<%'On Error Resume Next%>
<!--#include file="caminhos.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="funcoes_migra_notas.asp"-->
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

obr=co_materia&"$!$"&unidade&"$!$"&curso&"$!$"&etapa&"$!$"&turma&"$!$"&periodo&"$!$"&ano_letivo&"$!$"&co_prof

i=1

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_ns & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CONMT = Server.CreateObject("ADODB.Connection") 
		ABRIRMT = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONMT.Open ABRIRMT	
		
				
'		Set RSMT  = Server.CreateObject("ADODB.Recordset")
'		SQL_MT  = "Select CO_Materia_Principal from TB_Materia WHERE CO_Materia = '"& co_materia&"'"
'		Set RSMT  = CONMT.Execute(SQL_MT)
'		
'co_materia_pr = RSMT("CO_Materia_Principal")
'		
'if Isnull(co_materia_pr) then
'	co_materia_pr= co_materia
'else
'	co_materia_pr = co_materia_pr
'end if
			va_pt=request.form("pt")
			va_pp=request.form("pp")


	Set RS5a = Server.CreateObject("ADODB.Recordset")
	SQL5a = "SELECT * FROM TB_Programa_Subs where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia_Filha ='"& co_materia &"'"
	RS5a.Open SQL5a, CONMT
	
	co_materia_pr = RS5a("CO_Materia_Principal")		
	in_faltas = RS5a("IN_Faltas")
	in_bonus = RS5a("IN_Bonus")
	rec_semestral = RS5a("IN_Rec_Semestral")
	

fail = 0
for i=1 to max
	grava="ok"
	
	
	nu_matricula = request.form("nu_matricula_"&i)

	
	if nu_matricula = "falta" then
			i=i*1
			max=max*1
		if i=max then
		grava = "no"
		else
		   'Loop para descobrir o primeiro aluno após os faltantes
			while nu_matricula = "falta"
				i=i+1
				nu_matricula = request.form("nu_matricula_"&i)						
			wend	
				
			va_faltas=request.form("faltas_"&i)
			va_t1=request.form("t1_"&i)
			va_t2=request.form("t2_"&i)
			va_t3=request.form("t3_"&i)
			va_p1=request.form("p1_"&i)
			va_p2=request.form("p2_"&i)
			va_p3=request.form("p3_"&i)
			va_bon=request.form("bon_"&i)
			va_rec=request.form("rec_"&i)
		end if	
	else
			nu_matricula = request.form("nu_matricula_"&i)			
			va_faltas=request.form("faltas_"&i)
			va_t1=request.form("t1_"&i)
			va_t2=request.form("t2_"&i)
			va_t3=request.form("t3_"&i)
			va_p1=request.form("p1_"&i)
			va_p2=request.form("p2_"&i)
			va_p3=request.form("p3_"&i)
			va_bon=request.form("bon_"&i)
			va_rec=request.form("rec_"&i)
	end if
	
	if i=1 then
		vetor_matricula = nu_matricula	
	else
		vetor_matricula = vetor_matricula&", "&nu_matricula				
	end if	
	
if fail = 0 then 		
Session("va_faltas")=va_faltas
Session("va_pt")=va_pt
Session("va_pp")=va_pp
Session("va_t1")=va_t1
Session("va_t2")=va_t2
Session("va_t3")=va_t3
Session("va_p1")=va_p1
Session("va_p2")=va_p2
Session("va_p3")=va_p3
Session("va_bon")=va_bon
Session("va_rec")=va_rec	
end if

if not in_faltas then
	va_faltas=NULL
	
end if

if not in_bonus then
	va_bon = NULL
end if	

periodo = periodo*1
if (not rec_semestral) then
	va_rec = NULL	
end if	

'////////////////////////////////////////////////////////////////
'FALTAS
	if va_faltas="" or isnull(va_faltas) then
		if in_faltas then
			va_faltas=0	
		end if					
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
	
		if divisor_mt=0 THEN
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
	
		if divisor_mp=0 THEN
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

'	if media_t="ok" and media_p="ok" then
'		m1=((mt*va_pt)+(mp*va_pp))/(va_pt+va_pp)
'		'm1=m1*10
'		decimo = m1 - Int(m1)
'			If decimo >= 0.5 Then
'				nota_arredondada = Int(m1) + 1
'				m1=nota_arredondada
'			'elseif decimo > 0 Then
'			'	nota_arredondada = Int(m1) + 0.5
'			'	m1=nota_arredondada
'			else
'				nota_arredondada = Int(m1)
'				m1=nota_arredondada						
'			End If
'		'm1=m1/10							
'		m1 = formatNumber(m1,1)
'	else
'		m1=NULL
'	end if	
'		
'	if isnull(m1) or m1="" then
'		m2=NULL
'		m3=NULL	
'	else		
'		if isnull(va_bon) or va_bon="" then
'		m2=m1		
'		else
'			m1=m1*1		
'			va_bon=va_bon*1
'			m2=m1+va_bon
'			
'			if m2>100 then
'				if  fail = 1 then
'					grava = "no"
'				else											
'					fail = 1 
'					erro = "m2"
'					url = nu_matricula&"_"&va_bon&"_"&erro
'					grava = "no"
'				end if
'			end if
'			'm2=m2*10
'				decimo = m2 - Int(m2)
'					If decimo >= 0.5 Then
'						nota_arredondada = Int(m2) + 1
'						m2=nota_arredondada
'					'elseif decimo > 0 Then
'					'	nota_arredondada = Int(m2) + 0.5
'					'	m2=nota_arredondada
'					else
'						nota_arredondada = Int(m2)
'						m2=nota_arredondada											
'					End If
'			'm2=m2/10				
'				m2 = formatNumber(m2,1)
'		end if
'		if m2>10 then
'			m2=10
'		end if	
'		m3=m2
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
	'end if
grava_rec = "ok"	

if grava = "ok" then
	
'	response.Write(va_pt&"<-<br>")
'response.Write(va_pp&"<-<br>")
'response.Write(in_faltas&" -"&in_bonus)
'response.End()	



'response.Write(nu_matricula&"; "&co_materia_pr&"; "&co_materia&"; "&periodo&"; "&va_faltas&"; "&va_pt&"; "&va_pp&"; "&va_t1&"; "&va_t2&"; "&va_t3&"; "&va_p1&"; "&va_p2&"; "&va_p3&"; "&va_bon&"; "&va_rec&"; "&data&"; "&horario&"; "&co_usr)
'response.Write(grava&";"&fail&";"&url&"<-<br>")
'response.End()

if grava_rec = "ok" then
			Set RS0 = Server.CreateObject("ADODB.Recordset")
			CONEXAO0 = "Select * from TB_Nota_S WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
			Set RS0 = CON.Execute(CONEXAO0)
		If RS0.EOF THEN	
		
			Set RS = server.createobject("adodb.recordset")
			
			RS.open "TB_Nota_S", CON, 2, 2 'which table do you want open
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
				'RS("VA_Media1")=m1
				RS("VA_Bonus")=va_bon	
				'RS("VA_Media2")=m2
				RS("VA_Rec")=va_rec
				'RS("VA_Media3")=m3
				RS("DA_Ult_Acesso") = data
				RS("HO_ult_Acesso") = horario
				RS("CO_Usuario")= co_usr
			
			RS.update
			set RS=nothing
			
		else
	
			Set RS0 = Server.CreateObject("ADODB.Recordset")
			CONEXAO0 = "DELETE * from TB_Nota_S WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
			Set RS0 = CON.Execute(CONEXAO0)
	
			Set RS = server.createobject("adodb.recordset")
			
			RS.open "TB_Nota_S", CON, 2, 2 'which table do you want open
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
				'RS("VA_Media1")=m1
				RS("VA_Bonus")=va_bon	
				'RS("VA_Media2")=m2
				RS("VA_Rec")=va_rec
				'RS("VA_Media3")=m3
				RS("DA_Ult_Acesso") = data
				RS("HO_ult_Acesso") = horario
				RS("CO_Usuario")= co_usr
	
			
			RS.update
			set RS=nothing		
			
		end if
	end if		
end if

next




if fail = 1 then
	response.Redirect(endereco_origem&"notas.asp?or=01&opt=err6&hp=err_"&url&"&obr="&obr) 
else

	atualizou_mae = atualiza_disciplina_mae(vetor_matricula, curso, etapa, co_materia_pr, periodo, data, horario, co_usr)
	
	if atualizou_mae <> "S" then
		vetor_atualizou_mae = split(atualizou_mae, "$!$")
		materia_mae = vetor_atualizou_mae(0)		
		url = vetor_atualizou_mae(1)
		obr = obr&"&complemento="&materia_mae	
		response.Redirect(endereco_origem&"notas.asp?or=01&opt=err6&hp=err_"&url&"&obr="&obr) 				
	end if
END IF

outro="P:"&periodo&",D:"&co_materia&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma&""

			call GravaLog (chave,outro)

if opt="cln" then
outro="P:"&periodo&",D:"&co_materia&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma&", Comunicou"
			call GravaLog (chave,outro)
response.Redirect(endereco_origem&"comunicar.asp?or=01&nota=TB_Nota_S&opt=ok&obr="&obr)
comunicou = comunica_disc_mae(unidade, curso, etapa, co_prof, co_materia_pr, periodo, "TB_Nota_A")
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