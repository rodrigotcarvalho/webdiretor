<%On Error Resume Next%>
<!--#include file="caminhos.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="../../global/funcoes_diversas.asp"-->
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
			va_bon=request.form("bon_"&i)
			va_rec=request.form("rec_"&i)
	end if
	
if fail = 0 then 		
Session("faltas")=va_faltas
Session("t1")=va_t1
Session("t2")=va_t2
Session("t3")=va_t3
Session("p1")=va_p1
Session("p2")=va_p2
Session("bon")=va_bon
Session("rec")=va_rec	
end if

max_rec=10
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
								erro = "f$0"
								url = nu_matricula&"_"&va_faltas&"_"&erro
								grava = "no"
							end if
						end if
			else	
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "f$1"
					url = nu_matricula&"_"&va_faltas&"_"&erro
					grava = "no"
				end if
			end if		
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "f$0"
				url = nu_matricula&"_"&va_faltas&"_"&erro
				grava = "no"
			end if
		end if
	end if


'TESTES

if ano_letivo>= 2015 then
	s_va_t1=0
	if va_t1="" or isnull(va_t1) then
		va_t1=NULL		
		s_va_t1=0
		soma_va_t1=0		
	else
	    if co_materia <> "EDUF" and co_materia <> "EDUR" then
			teste_va_t1 = isnumeric(va_t1)
			if teste_va_t1= true then					
				va_t1=va_t1*1			
				if va_t1 =<10 then
						s_va_t1=1
						soma_va_t1=va_t1																										
				else
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "t1-a$2"
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
						erro = "t1-a$0"
						matric_Erro=i
						url = nu_matricula&"_"&va_t1&"_"&erro
						grava = "no"
				end if
			end if
		else
			va_t1=NULL	
		end if			
	end if
	
	if va_t2="" or isnull(va_t2) then
		va_t2=NULL		
		s_va_t2=0
		soma_va_t2=0
	else
	    if co_materia <> "EDUF" and co_materia <> "EDUR" then	
			teste_va_t2 = isnumeric(va_t2)		
			if teste_va_t2= true then							
				va_t2=va_t2*1	
				if va_t2 =<10 then 
						s_va_t2=1
						soma_va_t2=va_t2						
				else
					if  fail = 1 then
						grava = "no"
					else
						fail = 1
						erro = "t2-a$3"
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
					erro = "t2-a$0"
					matric_Erro=i
					url = nu_matricula&"_"&va_t2&"_"&erro
					grava = "no"
				end if
			end if
		else
			va_t2=NULL	
		end if				
	end if
	
	if va_t3="" or isnull(va_t3) then
		va_t3=NULL		
		s_va_t3=0
		soma_va_t3=0		
	else
	    if co_materia <> "EDUF" and co_materia <> "EDUR" then
			'if ano_letivo>= 2021 and (co_materia = "ART" or co_materia = "ARTE") then
			'	teste_va_t3 = isnumeric(va_t3)
			'	if teste_va_t3= true then							
			'		va_t3=va_t3*1
			'
			'		if va_t3 =<10 then
			'				s_va_t3=1
			'				soma_va_t3=va_t3																										
			'		else
			'			if  fail = 1 then
			'				grava = "no"
			'			else
			'				fail = 1 
			'				erro = "t3-a$3"
			'				matric_Erro=i
			'				url = nu_matricula&"_"&va_t3&"_"&erro
			'				grava = "no"
			'			end if					
			'		end if				
			'	else
			'		if  fail = 1 then
			'			grava = "no"
			'		else
			'			fail = 1 
			'			erro = "t3-a$0"
			'			matric_Erro=i
			'			url = nu_matricula&"_"&va_t3&"_"&erro
			'			grava = "no"
			'		end if
			'	end if
			'else
				if ano_letivo>= 2021 then
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "t3-a$4"
						matric_Erro=i
						url = nu_matricula&"_"&va_t3&"_"&erro
						grava = "no"
					end if				
				else
					teste_va_t3 = isnumeric(va_t3)
					if teste_va_t3= true then							
						va_t3=va_t3*1
			
						if va_t3 =<10 then
								s_va_t3=1
								soma_va_t3=va_t3																										
						else
							if  fail = 1 then
								grava = "no"
							else
								fail = 1 
								erro = "t3-a$3"
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
							erro = "t3-a$0"
							matric_Erro=i
							url = nu_matricula&"_"&va_t3&"_"&erro
							grava = "no"
						end if
					end if				
				end if
			'end if
		else
			va_t3=NULL	
		end if			
	end if

'//////////////////////////////////////////////////////////////////////
'Notas
	s_va_p=0

	if va_p1="" or isnull(va_p1) then
		va_p1=NULL		
		s_va_p1=0
		soma_va_p1=0
		limite_m1=10		
	else
	    if co_materia <> "EDUF" then	
			teste_va_p1 = isnumeric(va_p1)
			if teste_va_p1= true then							
				va_p1=va_p1*1	
				if va_p1 =<10 then
					s_va_p1=1
					soma_va_p1=va_p1																						
				else
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "p1-a$2"
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
						erro = "p1-a$0"
						matric_Erro=i
						url = nu_matricula&"_"&va_p1&"_"&erro
						grava = "no"
				end if
			end if
		else
			va_p1=NULL	
		end if			
	end if

	if va_p2="" or isnull(va_p2) then
		va_p2=NULL		
		s_va_p2=0
		soma_va_p2=0		
	else
	    'if co_materia = "ART" or co_materia = "ARTE" or co_materia = "EDUF" or co_materia = "EDUR" then
		if co_materia = "EDUF" or co_materia = "EDUR" then
			teste_va_p2 = isnumeric(va_p2)
			if teste_va_p2= true then	
				if va_p2 =<10 then
					s_va_p2=1
					soma_va_p2=va_p2																					
				else
					if  fail = 1 then
						grava = "no"
					else
							fail = 1 
							erro = "p2-a$2"
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
						erro = "p2-a$0"
						matric_Erro=i
						url = nu_matricula&"_"&va_p2&"_"&erro
						grava = "no"
				end if
			end if
		else
			va_p2=NULL	
		end if							
	end if



	if va_bon="" or isnull(va_bon) then
		va_bon=NULL		
		s_va_bon=0
		soma_va_bon=0
	else
		teste_va_bon = isnumeric(va_bon) 
		if teste_va_bon = true then
			if va_bon=<10 then
				va_bon=va_bon*1
				s_va_bon=1
				soma_va_bon=va_bon																				
			else
				if  fail = 1 then
					grava = "no"
				else											
					fail = 1 
					erro = "bon$1"
					url = nu_matricula&"_"&va_bon&"_"&erro
					grava = "no"
				end if			
			end if
		else
			if  fail = 1 then
				grava = "no"
			else											
				fail = 1 
				erro = "bon$0"
				url = nu_matricula&"_"&va_bon&"_"&erro
				grava = "no"
			end if
		end if
	end if
	
	if va_rec="" or isnull(va_rec) then
		va_rec=NULL		
		s_va_rec=0
		soma_va_rec=0		
	else
		
		teste_va_rec = isnumeric(va_rec) 
		if teste_va_rec = true then
			va_rec=va_rec*1
			if va_rec=<10 then
				va_rec=va_rec*1
				s_va_rec=1						
				soma_va_rec=va_rec						
								
			else
				if  fail = 1 then
					grava = "no"
				else					
					fail = 1 
					erro = "rec$1"
					url = nu_matricula&"_"&va_rec&"_"&erro
					grava = "no"
				end if							
			end if

		else
			if  fail = 1 then
				grava = "no"
			else					
				fail = 1 
				erro = "rec$0"
				url = nu_matricula&"_"&va_rec&"_"&erro
				grava = "no"
			end if
		end if
	end if	

'/////////////////////////////////////////////////////////////////////////
'Médias
	periodo=periodo*1
	'if periodo=4 then
'		str=NULL
'		if va_p1="" or isnull(va_p1) then
'			m1=NULL
'			m2=NULL
'			m3=NULL	
'		else	
'			m1=soma_va_p1
'			m2=soma_va_p1
'			m3=soma_va_p1		
''				decimo = m3 - Int(m3)
''				If decimo >= 0.75 Then
''					nota_arredondada = Int(m3) + 1
''					m3_arred=nota_arredondada
''				elseIf decimo >= 0.25 Then
''					nota_arredondada = Int(m3) + 0.5
''					m3_arred=nota_arredondada
''				else
''					nota_arredondada = Int(m3)
''					m3_arred=nota_arredondada											
''				End If			
''				m3 = formatNumber(m3_arred,1)	
'				m3=arredonda(m3,"mat_dez",1,outro)					
'		end if				
'	else
		s_va_t1=s_va_t1*1
		s_va_t2=s_va_t2*1
		s_va_t3=s_va_t3*1
		soma_va_t1=soma_va_t1*1
		soma_va_t2=soma_va_t2*1
		soma_va_t3=soma_va_t3*1	
		s_va_p1 = s_va_p1*1
		s_va_p2 = s_va_p2*1
		s_va_p3 = s_va_p3*1
		soma_va_p1 = soma_va_p1*1
		soma_va_p2 = soma_va_p2*1


		
		soma_simulados=NULL
		if periodo=4 then
			str=NULL
			if va_p1="" or isnull(va_p1) then
				m1=NULL
				m2=NULL
				m3=NULL	
			else	
				m1=soma_va_p1
				m2=soma_va_p1
				m3=soma_va_p1		
	'				decimo = m3 - Int(m3)
	'				If decimo >= 0.75 Then
	'					nota_arredondada = Int(m3) + 1
	'					m3_arred=nota_arredondada
	'				elseIf decimo >= 0.25 Then
	'					nota_arredondada = Int(m3) + 0.5
	'					m3_arred=nota_arredondada
	'				else
	'					nota_arredondada = Int(m3)
	'					m3_arred=nota_arredondada											
	'				End If			
	'				m3 = formatNumber(m3_arred,1)	
					m3=arredonda(m3,"mat_dez",1,outro)					
			end if				
		else
			'if co_materia = "ART" or co_materia = "ARTE" then
			'	if s_va_t1 = 1 and s_va_t2 = 1 and s_va_t3 = 1 and s_va_p1 = 1 and s_va_p2 = 1 then		
			'		m1=(soma_va_t1+soma_va_t2+soma_va_t3+(soma_va_p1*3)+(soma_va_p2*2))/8
			'	else
			'		m1=NULL		
			'	end if		  
			'elseif co_materia = "EDUR" then
			if co_materia = "EDUR" then
				if s_va_p1 = 1 and s_va_p2 = 1 then				
					 m1=(soma_va_p1+soma_va_p2)/2
				else
					m1=NULL		
				end if		  
			elseif co_materia = "EDUF" then
				if s_va_p2 = 1 then				
					m1=soma_va_p2
				else
					m1=NULL		
				end if		  
			else
				if s_va_t1 = 1 and s_va_t2 = 1 and s_va_p1 = 1 then
					if ano_letivo>= 2021 then
						m1=(soma_va_t1+soma_va_t2+(soma_va_p1*3))/5
					else
						if s_va_t3 = 1 then
							m1=(soma_va_t1+soma_va_t2+soma_va_t3+(soma_va_p1*3))/6
						else
							m1=NULL	
						end if
					end if
				else
					m1=NULL		
				end if
			end if	
		end if	
		if isnull(m1) or m1="" then
			m2=NULL
			m3=NULL	
		else	
			m1=arredonda(m1,"mat_dez",1,outro)	
				
			if isnull(va_bon) or va_bon="" then
				m2=m1	
				soma_bon = 0	
			else
				m1=m1*1		
				va_bon=va_bon*1
				m2=m1+va_bon
				soma_bon = va_bon					
				if m2>10 then
					m2=10
				end if
			end if	
			
			if isnull(va_rec) or va_rec="" then
				m3=m2
				if not isnull(m3) then		
					m3=arredonda(m3,"mat_dez",1,outro)							
				end if									
			else
			    if va_rec<soma_va_p1 then
					m3=m2					
				else	
					if co_materia = "ART" or co_materia = "ARTE" then
					'response.Write("("&soma_va_t1&"+"&soma_va_t2&"+"&soma_va_t3&"+("&va_rec&"*3)+("&soma_va_p2&"*2))/8) =  ")
							m3=(soma_va_t1+soma_va_t2+soma_va_t3+(va_rec*3)+(soma_va_p2*2))/8
						'response.Write(m3)	
										'response.End() 
					elseif co_materia = "EDUR" then
						m3=(va_rec+soma_va_p2)/2	  
					elseif co_materia = "EDUF" then		
						m3=va_rec	  
					else
						m3=(soma_va_t1+soma_va_t2+soma_va_t3+(va_rec*3))/6	
					end if				
					m3=m3+soma_bon
					m3=arredonda(m3,"mat_dez",1,outro)	
									
					if m3< m2 then
						m3=m2
					end if		
				end if			
			end if		
		end if
	'end if	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
'---------------------------------------------------------------------------------------------------------------------	
else
	s_va_t1=0
	if va_t1="" or isnull(va_t1) then
		va_t1=NULL		
		s_va_t1=0
		soma_va_t1=0		
	else
		teste_va_t1 = isnumeric(va_t1)
		if teste_va_t1= true then					
			va_t1=va_t1*1			
			if va_t1 =<1 then
					s_va_t1=1
					soma_va_t1=va_t1																										
			else
				if  fail = 1 then
					grava = "no"
				else
						fail = 1 
						erro = "t1-a$1"
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
					erro = "t1-a$0"
					matric_Erro=i
					url = nu_matricula&"_"&va_t1&"_"&erro
					grava = "no"
			end if
		end if
	end if

	curso = curso*1
	etapa = etapa*1		
	if co_materia = "ART" or co_materia = "ARTE" then
		limite_t2 = 1
	elseif (curso = 2 and etapa=1) then
		limite_t2 = 3	
		tipo_msg_erro="c2e1"	
	else
		limite_t2 = 2
		limite_t3 = 2																
	end if	
	
	
	if va_t2="" or isnull(va_t2) then
		va_t2=NULL		
		s_va_t2=0
		soma_va_t2=0
		limite_t3=4		
	else
		teste_va_t2 = isnumeric(va_t2)
		if teste_va_t2= true then					
			va_t2=va_t2*1	
			if va_t2 =<limite_t2 then 
					s_va_t2=1
					soma_va_t2=va_t2						
			else
				if  fail = 1 then
					grava = "no"
				else
						fail = 1
						if tipo_msg_erro="c2e1" then
							erro = "t2-a$2"						
						else 
							erro = "t2-a$1"
						end if	
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
					erro = "t2-a$0"
					matric_Erro=i
					url = nu_matricula&"_"&va_t2&"_"&erro
					grava = "no"
				end if
			end if
	end if
	curso=curso*1
	etapa=etapa*1
	if co_materia = "ART" or co_materia = "ARTE" then
		limite_t3 = 1
	elseif (curso = 2 and etapa=1) then
		limite_t3= 3 - soma_va_t2
		tipo_msg_erro="c2e1"			
	end if	
	
	if va_t3="" or isnull(va_t3) then
		va_t3=NULL		
		s_va_t3=0
		soma_va_t3=0		
	else
		teste_va_t3 = isnumeric(va_t3)
		if teste_va_t3= true then					
			va_t3=va_t3*1
			limite_t3=limite_t3*1			
			if va_t3 =<limite_t3 then
					s_va_t3=1
					soma_va_t3=va_t3																										
			else
				if  fail = 1 then
					grava = "no"
				else
						fail = 1 
						if tipo_msg_erro="c2e1" then
							erro = "t3-a$2"						
						else 
							erro = "t3-a$1"
						end if	
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
					erro = "t3-a$0"
					matric_Erro=i
					url = nu_matricula&"_"&va_t3&"_"&erro
					grava = "no"
			end if
		end if
	end if

'//////////////////////////////////////////////////////////////////////
'Notas
s_va_p=0
	curso = curso*1
	etapa = etapa*1		
	if curso = 2 and etapa=1 then
		limite_p1 = 6
		tipo_msg_erro="c2e1"		
	else
		limite_p1 = 5
	end if	

	if va_p1="" or isnull(va_p1) then
		va_p1=NULL		
		s_va_p1=0
		soma_va_p1=0
		limite_m1=10		
	else
		teste_va_p1 = isnumeric(va_p1)
		if teste_va_p1= true then					
			va_p1=va_p1*1	
				if periodo=4 then
					if va_p1 =<10 then
							s_va_p1=1
							soma_va_p1=va_p1	
							max_rec=5																						
					else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								erro = "p1-a$2"
								matric_Erro=i
								url = nu_matricula&"_"&va_p1&"_"&erro
								grava = "no"
						end if					
					end if					
				else				
					if va_p1 =<limite_p1 then
							s_va_p1=1
							soma_va_p1=va_p1
							limite_m1=5	
							max_rec=5																								
					else
						if  fail = 1 then
							grava = "no"
						else
								fail = 1 
								if tipo_msg_erro="c2e1" then
									erro = "p1-a$3"						
								else 
									erro = "p1-a$1"
								end if									
								matric_Erro=i
								url = nu_matricula&"_"&va_p1&"_"&erro
								grava = "no"
						end if					
					end if	
				end if				
			else
				if  fail = 1 then
					grava = "no"
				else
						fail = 1 
						erro = "p1-a$0"
						matric_Erro=i
						url = nu_matricula&"_"&va_p1&"_"&erro
						grava = "no"
				end if
			end if
	end if
	if co_materia = "ART" or co_materia = "ARTE" then
		limite_p2 = 2
	else
		limite_p2 = limite_m1	
	end if
	if va_p2="" or isnull(va_p2) then
		va_p2=NULL		
		s_va_p2=0
		soma_va_p2=0		
	else
		teste_va_p2 = isnumeric(va_p2)
		if teste_va_p2= true then					
		va_p2=va_p2*1			
				if va_p2 =<limite_p2 then
						s_va_p2=1
						soma_va_p2=va_p2																					
				else
					if  fail = 1 then
						grava = "no"
					else
							fail = 1 
							erro = "p2-a$1"
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
						erro = "p2-a$0"
						matric_Erro=i
						url = nu_matricula&"_"&va_p2&"_"&erro
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
		if teste_va_bon = true then
			if va_bon=<10 then
				va_bon=va_bon*1
				s_va_bon=1
				soma_va_bon=va_bon																				
			else
				if  fail = 1 then
					grava = "no"
				else											
					fail = 1 
					erro = "bon$1"
					url = nu_matricula&"_"&va_bon&"_"&erro
					grava = "no"
				end if			
			end if
		else
			if  fail = 1 then
				grava = "no"
			else											
				fail = 1 
				erro = "bon$0"
				url = nu_matricula&"_"&va_bon&"_"&erro
				grava = "no"
			end if
		end if
	end if
	
	curso = curso*1
	etapa = etapa*1		
	if curso = 2 and etapa=1 then
		max_rec = 6
	end if		

	if va_rec="" or isnull(va_rec) then
		va_rec=NULL		
		s_va_rec=0
		soma_va_rec=0		
	else
		
		teste_va_rec = isnumeric(va_rec) 
		if teste_va_rec = true then
			max_rec=max_rec*1
			va_rec=va_rec*1
			if va_rec=<max_rec then
				va_rec=va_rec*1
				s_va_rec=1						
				soma_va_rec=va_rec						
								
			else
				if  fail = 1 then
					grava = "no"
				else					
					fail = 1 
					erro = "rec$1"
					url = nu_matricula&"_"&va_rec&"_"&erro
					grava = "no"
				end if							
			end if

		else
			if  fail = 1 then
				grava = "no"
			else					
				fail = 1 
				erro = "rec$0"
				url = nu_matricula&"_"&va_rec&"_"&erro
				grava = "no"
			end if
		end if
	end if	

'/////////////////////////////////////////////////////////////////////////
'Médias
	periodo=periodo*1
	if periodo=4 then
		str=NULL
		if va_p1="" or isnull(va_p1) then
			m1=NULL
			m2=NULL
			m3=NULL	
		else	
			m1=soma_va_p1
			m2=soma_va_p1
			m3=soma_va_p1		
'				decimo = m3 - Int(m3)
'				If decimo >= 0.75 Then
'					nota_arredondada = Int(m3) + 1
'					m3_arred=nota_arredondada
'				elseIf decimo >= 0.25 Then
'					nota_arredondada = Int(m3) + 0.5
'					m3_arred=nota_arredondada
'				else
'					nota_arredondada = Int(m3)
'					m3_arred=nota_arredondada											
'				End If			
'				m3 = formatNumber(m3_arred,1)	
				m3=arredonda(m3,"mat_dez",1,outro)					
		end if				
	else
		s_va_t1=s_va_t1*1
		s_va_t2=s_va_t2*1
		s_va_t3=s_va_t3*1
		soma_va_t1=soma_va_t1*1
		soma_va_t2=soma_va_t2*1
		soma_va_t3=soma_va_t3*1	
		s_va_p1 = s_va_p1*1
		s_va_p2 = s_va_p2*1
		s_va_p3 = s_va_p3*1
		soma_va_p1 = soma_va_p1*1
		soma_va_p2 = soma_va_p2*1

		if co_materia = "ART" or co_materia = "ARTE" then
		  m1=soma_va_t1+soma_va_t2+soma_va_t3+soma_va_p1+soma_va_p2
		else		
		
			soma_simulados=soma_va_t2+soma_va_t3
			
			if s_va_t2=0 and s_va_t3=0 and soma_simulados=0 then
				s_va_ss=0
				soma_simulados=NULL
				soma_simulados_m1=0
			else
				s_va_ss=1
				soma_simulados_m1=soma_simulados
			end if
		

			
			divisor_mp=s_va_p1+s_va_p2+s_va_p3
			dividendo_mp=soma_va_p1+soma_va_p2
		'	response.Write(s_va_t1&"-"&s_va_t2&"-"&s_va_t3&"-"&soma_va_p1&"-"&s_va_p2&"-"&soma_va_p2&"<BR>")	
		
	
		
			if s_va_p2=1 then
			
				m1=soma_va_p1+soma_va_p2
		
		'response.Write(soma_va_p1&"-"&soma_va_t1&"-"&soma_simulados&"-"&soma_va_p2&"<BR>")	
		'if i=3 then	
		
		'	response.end()
		
		'end if
				
					if m1>10 then
								if  fail = 1 then
									grava = "no"
								else					
									fail = 1 
									erro = "m1"
									url = nu_matricula&"_"&m1&"_"&erro
									grava = "no"
								end if							
					end if			
				elseif s_va_t1=1 and s_va_ss=1 and s_va_p1=1 then
					m1=soma_va_p1+soma_va_t1+soma_simulados
						if m1>10 then
								if  fail = 1 then
									grava = "no"
								else					
									fail = 1 
									erro = "m1"
									url = nu_matricula&"_"&m1&"_"&erro
									grava = "no"
								end if							
						end if			
			else
			m1=NULL		
			end if
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
				
				if m2>10 then
					if  fail = 1 then
						grava = "no"
					else											
						fail = 1 
						erro = "m2"
						url = nu_matricula&"_"&va_bon&"_"&erro
						grava = "no"
					end if
				end if
			end if
			
			'if criado por conta do e-mail de 27/05/2013 11:08 AM
			'if ano_letivo >= 2013 then	
				
		
			if ano_letivo >= 2013 then									
				if isnull(va_rec) or va_rec="" then
							m3=m2
					if isnull(m3) then
					else			
						m3=arredonda(m3,"mat_dez",1,outro)							
					end if									
				else
	  				'if criado por conta dos e-mails de 21/06/2013 02:32 PM	e 27/06/2013 8:48 AM			
					if co_materia = "ART" or co_materia = "ARTE" then
						IF s_va_p1=1 then
							if va_rec>soma_va_p1 then
									m3=m2-soma_va_p1+soma_va_rec
								else
									m3=m2
							end if
						else
							m3=m2
						end if									
					else
						IF s_va_p1=1 AND s_va_p2=1 THEN
							IF soma_va_p1>soma_va_p2 THEN
								rec_compara = soma_va_p2
								soma_com_rec=soma_va_p1
							else
								rec_compara = soma_va_p1
								soma_com_rec=soma_va_p2
							end if
							
																
							va_rec=va_rec*1
	
							if va_rec>rec_compara then
									m3=soma_com_rec+soma_va_rec+soma_va_bon
								else
									m3=m2
							end if					
						ELSEIF s_va_p1=1 AND s_va_p2=0 THEN
							va_rec=va_rec*1
							soma_va_p1=soma_va_p1*1					
							if s_va_ss=1 then
								ss_rec = va_rec*0.8			
							end if	
												
								soma_va_rec = soma_va_rec*1
								soma_va_p1 = soma_va_p1*1
								if va_rec>soma_va_p1 then
									m3a=soma_va_t1+soma_simulados+soma_va_rec+soma_va_bon
								else
									m3a=m2
								end if	
								ss_rec=ss_rec*1
								soma_simulados = soma_simulados*1
								if ss_rec>soma_simulados then
									m3b=soma_va_t1+ss_rec+soma_va_p1+soma_va_bon
								else
									m3b=m2
								end if
								
								if m3a>	m3b then
									m3 = m3a
								else
									m3 = m3b							
								end if	
															
						ELSEIF s_va_p1=0 AND s_va_p2=1 THEN
							va_rec=va_rec*1
							soma_va_p2=soma_va_p2*1		
							if s_va_ss=1 then
								soma_simulados_rec = soma_simulados_m1*0.8		
							end if	
																	
							if va_rec>soma_va_p2 then
								m3=va_rec
							else
								m3=m2
							end if	
						END IF
						m3=arredonda(m3,"mat_dez",1,outro)		
					end if
				end if	
			else	
				if isnull(va_rec) or va_rec="" then
							m3=m2
					if isnull(m3) then
					else			
	'					decimo = m3 - Int(m3)
	'					If decimo >= 0.75 Then
	'						nota_arredondada = Int(m3) + 1
	'						m3=nota_arredondada
	'					elseIf decimo >= 0.25 Then
	'						nota_arredondada = Int(m3) + 0.5
	'						m3=nota_arredondada
	'					else
	'						nota_arredondada = Int(m3)
	'						m3=nota_arredondada											
	'					End If			
	'					m3 = formatNumber(m3,1)		
						m3=arredonda(m3,"mat_dez",1,outro)							
					end if									
				else
					if co_materia_pr= "EDUR" then
						IF s_va_p1=1 AND s_va_p2=1 THEN
							IF soma_va_p1>soma_va_p2 THEN
							rec_compara = soma_va_p2
							soma_com_rec=soma_va_p1
							else
							rec_compara = soma_va_p1
							soma_com_rec=soma_va_p2
							end if
			
							va_rec=va_rec*1
							rec_compara=rec_compara*1
							if va_rec>rec_compara then
									m3=soma_com_rec+soma_va_rec+soma_va_bon
								else
									m3=m2
							end if
						end if	
					else	
						IF s_va_p1=1 AND s_va_p2=1 THEN
							IF soma_va_p1>soma_va_p2 THEN
							rec_compara = soma_va_p2
							soma_com_rec=soma_va_p1
							else
							rec_compara = soma_va_p1
							soma_com_rec=soma_va_p2
							end if
			
							va_rec=va_rec*1
							rec_compara=rec_compara*1
							if va_rec>rec_compara then
									m3=soma_com_rec+soma_va_rec+soma_va_bon
								else
									m3=m2
							end if					
						ELSEIF s_va_p1=1 AND s_va_p2=0 THEN
							va_rec=va_rec*1
							soma_va_p1=soma_va_p1*1
								if va_rec>soma_va_p1 then
									m3=soma_va_t1+soma_simulados+soma_va_rec+soma_va_bon
								else
									m3=m2
								end if	
						ELSEIF s_va_p1=0 AND s_va_p2=1 THEN
							va_rec=va_rec*1
							soma_va_p2=soma_va_p2*1			
								if va_rec>soma_va_p2 then
									m3=va_rec
								else
									m3=m2
								end if	
						END IF
								
	'					decimo = m3 - Int(m3)
	'					If decimo >= 0.75 Then
	'						nota_arredondada = Int(m3) + 1
	'						m3=nota_arredondada
	'					elseIf decimo >= 0.25 Then
	'						nota_arredondada = Int(m3) + 0.5
	'						m3=nota_arredondada
	'					else
	'						nota_arredondada = Int(m3)
	'						m3=nota_arredondada											
	'					End If	
	'					m3 = formatNumber(m3,1)					
						m3=arredonda(m3,"mat_dez",1,outro)		
	
					end if
				end if
			end if				
		end if
	end if	
end if		

if isnumeric(m3) and m3>10 then
	m3=10
end if	

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
			RS("VA_TR")=va_t1
			RS("VA_S1")=va_t2
			RS("VA_S2")=va_t3
			RS("VA_SS")=soma_simulados
			RS("VA_PR")=va_p1
			RS("VA_ATV")=va_p2					
			RS("VA_Media1")=m1
			RS("VA_Extra")=va_bon	
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
			RS("VA_TR")=va_t1
			RS("VA_S1")=va_t2
			RS("VA_S2")=va_t3
			RS("VA_SS")=soma_simulados
			RS("VA_PR")=va_p1
			RS("VA_ATV")=va_p2					
			RS("VA_Media1")=m1
			RS("VA_Extra")=va_bon	
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

next
if fail = 1 then

response.Redirect(endereco_origem&"notas.asp?or=01&opt=err6&obr="&obr&"&hp=err_"&url) 

END IF

outro="P:"&periodo&",D:"&co_materia&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma&""

			call GravaLog (chave,outro)

if opt="cln" then
outro="P:"&periodo&",D:"&co_materia&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma&", Comunicou"
			call GravaLog (chave,outro)
response.Redirect(endereco_origem&"comunicar.asp?or=01&nota=TB_Nota_B&opt=ok&obr="&obr)
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