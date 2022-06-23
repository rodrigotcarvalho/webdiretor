<%On Error Resume Next%>
<!--#include file="../../global/funcoes_diversas.asp" -->
<!--#include file="caminhos.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="funcoes6.asp"-->

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

	tp_modelo=tipo_divisao_ano(curso,etapa,"tp_modelo")


if co_materia="EDFS" or co_materia="EDAR" or co_materia="INFO" or co_materia="EDRE" then
peso_t1=100
peso_t2=100
peso_t3=100
peso_p1=100
peso_p2=100
peso_p3=100
peso_bon=100
peso_rec=100
peso_mt=1 
peso_mp=1
else
peso_t1=100
peso_t2=100
peso_t3=100
peso_p1=100
peso_p2=100
peso_p3=100
peso_bon=100
peso_rec=100
peso_mt=1 
peso_mp=1
end if

obr=co_materia&"$!$"&unidade&"$!$"&curso&"$!$"&etapa&"$!$"&turma&"$!$"&periodo&"$!$"&ano_letivo&"$!$"&co_prof

i=1

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_na & ";Driver={Microsoft Access Driver (*.mdb)}"
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

fail = 0
for i=1 to max
	grava="ok"
	
	
	nu_matricula = request.form("nu_matricula_"&i)
	
		'	response.Write(i&" - "&nu_matricula&"<<BR>")
	
	if nu_matricula = "falta" then
			i=i*1
			max=max*1
		if i=max then
		grava = "no"
		else
			while nu_matricula = "falta"
			i=i+1
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
			wend
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

	
if fail = 0 then 		
Session("va_faltas")=va_faltas
Session("va_t1")=va_t1
Session("va_t2")=va_t2
Session("va_t3")=va_t3
Session("va_p1")=va_p1
Session("va_p2")=va_p2
Session("va_p3")=va_p3
Session("va_bon")=va_bon
Session("va_rec")=va_rec	
end if


'////////////////////////////////////////////////////////////////
'FALTAS
	if va_faltas="" or isnull(va_faltas) then
		va_faltas=NULL			
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



'TESTES
s_va_t=0
	if va_t1="" or isnull(va_t1) then
		va_t1=NULL		
		s_va_t1=0
		soma_va_t1=0		
	else
		teste_va_t1 = isnumeric(va_t1)
		if teste_va_t1= true then					
		va_t1=va_t1*1
		peso_t1=peso_t1*1			
					if va_t1 =<peso_t1 then
						'IF Int(va_av1)=va_av1 THEN
							s_va_t1=1
							soma_va_t1=va_t1						
						'ELSE	
						'	if  fail = 1 then
						'		grava = "no"
						'	else
						'		fail = 1 
						'		erro = "a1"
						'		matric_Erro=i
						'		url = nu_matricula&"_"&va_av1&"_"&erro
						'		grava = "no"
						'	end if					
						'end if																				
					else
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
								url = nu_matricula&"_"&va_av1&"_"&erro
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
		peso_t2=peso_t2*1			
					if va_t2 =<peso_t2 then
						'IF Int(va_av1)=va_av1 THEN
							s_va_t2=1
							soma_va_t2=va_t2						
						'ELSE	
						'	if  fail = 1 then
						'		grava = "no"
						'	else
						'		fail = 1 
						'		erro = "a1"
						'		matric_Erro=i
						'		url = nu_matricula&"_"&va_av1&"_"&erro
						'		grava = "no"
						'	end if					
						'end if																				
					else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "t2"
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
								erro = "t2"
								matric_Erro=i
								url = nu_matricula&"_"&va_av1&"_"&erro
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
		peso_t3=peso_t3*1			
					if va_t3 =<peso_t3 then
						'IF Int(va_av1)=va_av1 THEN
							s_va_t3=1
							soma_va_t3=va_t3						
						'ELSE	
						'	if  fail = 1 then
						'		grava = "no"
						'	else
						'		fail = 1 
						'		erro = "a1"
						'		matric_Erro=i
						'		url = nu_matricula&"_"&va_av1&"_"&erro
						'		grava = "no"
						'	end if					
						'end if																				
					else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "t3"
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
								erro = "t3"
								matric_Erro=i
								url = nu_matricula&"_"&va_av1&"_"&erro
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
		peso_p1=peso_p1*1			
					if va_p1 =<peso_p1 then
						'IF Int(va_av1)=va_av1 THEN
							s_va_p1=1
							soma_va_p1=va_p1						
						'ELSE	
						'	if  fail = 1 then
						'		grava = "no"
						'	else
						'		fail = 1 
						'		erro = "a1"
						'		matric_Erro=i
						'		url = nu_matricula&"_"&va_av1&"_"&erro
						'		grava = "no"
						'	end if					
						'end if																				
					else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "p1"
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
								erro = "p1"
								matric_Erro=i
								url = nu_matricula&"_"&va_av1&"_"&erro
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
		peso_p2=peso_p2*1			
					if va_p2 =<peso_p2 then
						'IF Int(va_av1)=va_av1 THEN
							s_va_p2=1
							soma_va_p2=va_p2						
						'ELSE	
						'	if  fail = 1 then
						'		grava = "no"
						'	else
						'		fail = 1 
						'		erro = "a1"
						'		matric_Erro=i
						'		url = nu_matricula&"_"&va_av1&"_"&erro
						'		grava = "no"
						'	end if					
						'end if																				
					else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "p2"
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
								erro = "p2"
								matric_Erro=i
								url = nu_matricula&"_"&va_av1&"_"&erro
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
		peso_p3=peso_p3*1			
					if va_p3 =<peso_p3 then
						'IF Int(va_av1)=va_av1 THEN
							s_va_p3=1
							soma_va_p3=va_p3						
						'ELSE	
						'	if  fail = 1 then
						'		grava = "no"
						'	else
						'		fail = 1 
						'		erro = "a1"
						'		matric_Erro=i
						'		url = nu_matricula&"_"&va_av1&"_"&erro
						'		grava = "no"
						'	end if					
						'end if																				
					else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "p3"
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
								erro = "p3"
								matric_Erro=i
								url = nu_matricula&"_"&va_av1&"_"&erro
								grava = "no"
						end if
			end if
	end if
	
	if va_bon="" or isnull(va_bon) then
		va_bon=NULL		
		s_va_bon=0
		soma_va_bon=0		
	else
		teste_va_bon = isnumeric(va_bon)
		if teste_va_bon= true then					
		va_bon=va_bon*1
		peso_bon=peso_bon*1			
					if va_bon =<peso_bon then
						'IF Int(va_av1)=va_av1 THEN
							s_va_bon=1
							soma_va_bon=va_bon						
						'ELSE	
						'	if  fail = 1 then
						'		grava = "no"
						'	else
						'		fail = 1 
						'		erro = "a1"
						'		matric_Erro=i
						'		url = nu_matricula&"_"&va_av1&"_"&erro
						'		grava = "no"
						'	end if					
						'end if																				
					else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "ext"
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
								erro = "ext"
								matric_Erro=i
								url = nu_matricula&"_"&va_av1&"_"&erro
								grava = "no"
						end if
			end if
	end if
	


	if va_rec="" or isnull(va_rec) then
		va_rec=NULL		
		s_va_rec=0	
		soma_va_rec = 0	
	else
		teste_va_rec = isnumeric(va_rec) 
		if teste_va_rec = true then
				va_rec= va_rec*1
				peso_rec = peso_rec*1
			if va_rec=<peso_rec then			
						va_rec=va_rec*1
						'IF Int(va_rec)=va_rec THEN
							s_va_rec = 1
							soma_va_rec=va_rec						
						'ELSE	
						'	if  fail = 1 then
						'		grava = "no"
						'	else					
						'		fail = 1 
						'		erro = "rec"
						'		url = nu_matricula&"_"&va_rec&"_"&erro
						'		grava = "no"
						'	end if					
						'end if								
			else
						if  fail = 1 then
							grava = "no"
						else					
							fail = 1 
							erro = "recp"
							url = nu_matricula&"_"&va_rec&"_"&erro
							grava = "no"
						end if							
			end if

		else
						if  fail = 1 then
							grava = "no"
						else					
							fail = 1 
							erro = "recn"
							url = nu_matricula&"_"&va_rec&"_"&erro
							grava = "no"
						end if
		end if
	end if	

'/////////////////////////////////////////////////////////////////////////
'Médias

	'Media teste
	denominador_teste = s_va_t1+s_va_t2+s_va_t3
	if denominador_teste>0 then
		va_mt =  (soma_va_t1+soma_va_t2+soma_va_t3)/denominador_teste
		va_mt =arredonda(va_mt,"mat",0,outro)
	else
		va_mt = NULL		
	end if
	
	'Media prova
	denominador_prova = s_va_p1+s_va_p2+s_va_p3
	if denominador_prova>0 then
		va_mp =  (soma_va_p1+soma_va_p2+soma_va_p3)/denominador_prova
		va_mp =arredonda(va_mp,"mat",0,outro)		
	else
		va_mp = NULL		
	end if
	
	'Media 1
	if denominador_teste>0 and denominador_prova>0 then
		peso_mt = peso_mt*1
		peso_mp = peso_mp*1
		media_1 = ((va_mt*peso_mt)+(va_mp*peso_mp))/(peso_mt+peso_mp)
		media_1 =arredonda(media_1,"mat",0,outro)			
		media_1 = media_1/10
		media_1 =arredonda(media_1,"quarto",1,outro)	
		media_1 = media_1*10				
	else
		media_1 = NULL
	end if

	'Media 2
	if denominador_teste>0 and denominador_prova>0 then
		media_1 = media_1*1
		soma_va_bon = soma_va_bon*1
		media_2 = media_1+soma_va_bon
		if media_2>100 then
			media_2 = 100
		end if	
	else
		media_2 = media_1
	end if
	
	'Media 3
	if denominador_teste>0 and denominador_prova>0 and s_va_rec = 1 then
		soma_va_rec = soma_va_rec*1
		media_2 = media_2*1
		if soma_va_rec>media_2 then
			media_3=soma_va_rec
		else
			media_3=media_2
		end if
		
		if media_3>100 then
			media_3 = 100
		end if	
	else
		media_3 = media_2
	end if	


'	response.Write(periodo&"-"&tp_modelo)
'		if s_va_t=0 and ((periodo<5 and tp_modelo="B") or (periodo<4 and tp_modelo="T")) THEN
'			media_1="no"
'			m1=NULL						
'		else
'			media_1="ok"
'			periodo=periodo*1
'			if ((periodo=5 and tp_modelo="B") or (periodo=4 and tp_modelo="T"))  then
'					if va_rf="" or isnull(va_rf) then
'						va_rf=NULL		
'						s_va_rf=0
'						media_1="no"
'						m1=NULL						
'					else
'						teste_va_rf = isnumeric(va_rf) 
'						if teste_va_rf = true then
'							if va_rf=<10 then
'									va_rf=va_rf*1
'										'IF Int(va_rf)=va_rf THEN
'											s_va_rf=va_rf
'											m1=s_va_rf																								
'										'ELSE						
'										'	if  fail = 1 then
'										'		grava = "no"
'										'	else												
'										'		fail = 1 
'										'		erro = "rf"
'										'		url = nu_matricula&"_"&va_rf&"_"&erro
'										'		grava = "no"
'										'	end if					
'										'end if
'										m1 = formatNumber(m1,1)																		
'							else
'										if  fail = 1 then
'											grava = "no"
'										else											
'											fail = 1 
'											erro = "rf"
'											url = nu_matricula&"_"&va_rf&"_"&erro
'											grava = "no"
'										end if			
'							end if
'				
'						else
'										if  fail = 1 then
'											grava = "no"
'										else											
'											fail = 1 
'											erro = "rf"
'											url = nu_matricula&"_"&va_rf&"_"&erro
'											grava = "no"
'										end if
'						end if
'					end if			
'			
'			elseif ((periodo=6 and tp_modelo="B") or (periodo=5 and tp_modelo="T")) then
'			if va_pf="" or isnull(va_pf) then
'						va_rf=NULL		
'						s_va_pf=0
'						media_1="no"
'						m1=NULL						
'					else
'						teste_va_pf = isnumeric(va_pf) 
'						if teste_va_pf = true then
'							if va_pf=<10 then
'									va_pf=va_pf*1
'										'IF Int(va_pf)=va_pf THEN
'											s_va_pf=va_pf
'											m1=s_va_pf																								
'										'ELSE						
'										'	if  fail = 1 then
'										'		grava = "no"
'										'	else												
'										'		fail = 1 
'										'		erro = "pf"
'										'		url = nu_matricula&"_"&va_pf&"_"&erro
'										'		grava = "no"
'										'	end if					
'										'end if
'										m1 = formatNumber(m1,1)																		
'							else
'										if  fail = 1 then
'											grava = "no"
'										else											
'											fail = 1 
'											erro = "pf"
'											url = nu_matricula&"_"&va_pf&"_"&erro
'											grava = "no"
'										end if			
'							end if
'				
'						else
'										if  fail = 1 then
'											grava = "no"
'										else											
'											fail = 1 
'											erro = "pf"
'											url = nu_matricula&"_"&va_pf&"_"&erro
'											grava = "no"
'										end if
'						end if
'					end if
'			
'			
'			
'			
'			else		
'				if co_materia="EDFS" or co_materia="EDAR" or co_materia="INFO" or co_materia="EDRE" then
'							m1=va_av1+va_av2
'				else				
'							m1=va_av1+va_av2+va_for
'				end if
'
'				'm1=m1*10
'					'decimo = m1 - Int(m1)
'						'If decimo >= 0.5 Then
'						'	nota_arredondada = Int(m1) + 1
'						'	m1=nota_arredondada
'
'						'elseif decimo > 0 Then
'						'	nota_arredondada = Int(m1) + 0.5
'						'	m1=nota_arredondada
'					'	if isnull(m1) or m1="" then
'		m2=NULL
'		m3=NULL	
'	else		
'		if isnull(va_ext) or va_ext="" then
'		m2=m1		
'		else
'			m1=m1*1		
'			va_ext=va_ext*1
'			m2=m1+va_ext
'			
'			if m2>10 then
'				if  fail = 1 then
'					grava = "no"
'				else											
'					fail = 1 
'					erro = "m2"
'					url = nu_matricula&"_"&va_ext&"_"&erro
'					grava = "no"
'				end if
'			end if
'			'm2=m2*10
'				'decimo = m2 - Int(m2)
'					'If decimo >= 0.5 Then
'					'	nota_arredondada = Int(m2) + 1
'					'	m2=nota_arredondada
'					'elseif decimo > 0 Then
'					'	nota_arredondada = Int(m2) + 0.5
'					'	m2=nota_arredondada
'					'else
'					'	nota_arredondada = Int(m2)
'					'	m2=nota_arredondada											
'					'End If
'			'm2=m2/10				
'				m2 = formatNumber(m2,1)
'		end if
'		if ((periodo<5 and tp_modelo="B") or (periodo<4 and tp_modelo="T")) then		
'			if isnull(va_rec) or va_rec="" then
'					decimo = m2 - Int(m2)
'					If decimo >= 0.75 Then
'						nota_arredondada = Int(m2) + 1
'						m2_arred=nota_arredondada				
'					elseIf decimo >= 0.25 Then
'						nota_arredondada = Int(m2) + 0.5
'						m2_arred=nota_arredondada
'					else
'						nota_arredondada = Int(m2)
'						m2_arred=nota_arredondada											
'					End If			
'					m2_arred = formatNumber(m2_arred,1)
'					m3=m2_arred						
'			else
'				if periodo = 1 or periodo = 2 then					
'						m2=m2*1
'						va_rec=va_rec*1
'						m3_temp=m2
'						decimo = m3_temp - Int(m3_temp)
'							If decimo >= 0.75 Then
'								nota_arredondada = Int(m3_temp) + 1
'								m3_temp=nota_arredondada
'							elseIf decimo >= 0.25 Then
'								nota_arredondada = Int(m3_temp) + 0.5
'								m3_temp=nota_arredondada
'							else
'								nota_arredondada = Int(m3_temp)
'								m3_temp=nota_arredondada											
'							End If			
'							m3_temp = formatNumber(m3_temp,1)
'						m2=m2*1
'						m3_temp=m3_temp*1						
'						if m3_temp >= m2 then
'							m3=m3_temp
'						else
'							decimo = m2 - Int(m2)
'							If decimo >= 0.75 Then
'								nota_arredondada = Int(m2) + 1
'								m2_arred=nota_arredondada
'							elseIf decimo >= 0.25 Then
'								nota_arredondada = Int(m2) + 0.5
'								m2_arred=nota_arredondada
'							else
'								nota_arredondada = Int(m2)
'								m2_arred=nota_arredondada											
'							End If			
'							m2_arred = formatNumber(m2_arred,1)
'							m3=m2_arred
'						end if
'				else
'						decimo = m2 - Int(m2)
'						If decimo >= 0.75 Then
'							nota_arredondada = Int(m2) + 1
'							m2_arred=nota_arredondada
'						elseIf decimo >= 0.25 Then
'							nota_arredondada = Int(m2) + 0.5
'							m2_arred=nota_arredondada
'						else
'							nota_arredondada = Int(m2)
'							m2_arred=nota_arredondada											
'						End If			
'						m2_arred = formatNumber(m2_arred,1)
'						m3=m2_arred
'				end if	
'			end if				
'		else
'			m3=m2	
'		end if										
'	end if	'else
'						'	nota_arredondada = Int(m1)
'						'	m1=nota_arredondada						
'						'End If
'					'm1=m1/10				
'					m1 = formatNumber(m1,1)					
'				end if
'		end if
			

if grava = "ok" then
	
		'	response.Write("Select * from TB_Nota_A WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo)

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		CONEXAO0 = "Select * from TB_Nota_A WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
		Set RS0 = CON.Execute(CONEXAO0)
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
			RS("MD_Teste")=va_mt
			RS("PE_Teste")=peso_mt
			RS("VA_Prova1")=va_p1
			RS("VA_Prova2")=va_p2
			RS("VA_Prova3")=va_p3
			RS("MD_Prova")=va_mp
			RS("PE_Prova")=peso_mp		
			RS("VA_Media1")=media_1
			RS("VA_Bonus")=va_bon	
			RS("VA_Media2")=media_2
			RS("VA_Rec")=va_rec
			RS("VA_Media3")=media_3
			RS("DA_Ult_Acesso") = data
			RS("HO_ult_Acesso") = horario
			RS("CO_Usuario")= co_usr
		
		RS.update
		set RS=nothing
		
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
			RS("MD_Teste")=va_mt
			RS("PE_Teste")=peso_mt
			RS("VA_Prova1")=va_p1
			RS("VA_Prova2")=va_p2
			RS("VA_Prova3")=va_p3
			RS("MD_Prova")=va_mp
			RS("PE_Prova")=peso_mp		
			RS("VA_Media1")=media_1
			RS("VA_Bonus")=va_bon	
			RS("VA_Media2")=media_2
			RS("VA_Rec")=va_rec
			RS("VA_Media3")=media_3
			RS("DA_Ult_Acesso") = data
			RS("HO_ult_Acesso") = horario
			RS("CO_Usuario")= co_usr
		
		RS.update
		set RS=nothing		
		
	end if
end if

next
if fail = 1 then

response.Redirect(endereco_origem&"notas.asp?or=01&opt=err6&obr="&obr&"&hp=err_"&url) 

END IF

outro="P:"&periodo&",D:"&co_materia&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma&""

			call GravaLog (chave,outro)

if opt="cln" then
outro=outro&", Comunicou"
			call GravaLog (chave,outro)
response.Redirect(endereco_origem&"comunicar.asp?or=01&nota=TB_Nota_A&opt=ok&obr="&obr)
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
response.redirect("../../../../inc/erro.asp")
end if
%>