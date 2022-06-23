<%On Error Resume Next%>
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

obr=co_materia&"$!$"&unidade&"$!$"&curso&"$!$"&etapa&"$!$"&turma&"$!$"&periodo&"$!$"&ano_letivo&"$!$"&co_prof

i=1

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_nb & ";Driver={Microsoft Access Driver (*.mdb)}"
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
			va_pp1=request.form("pp1")
			va_pp2=request.form("pp2")


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
			wend			
				va_faltas=request.form("faltas_"&i)						
				va_t1=request.form("t1_"&i)
				va_t2=request.form("t2_"&i)
				va_p1=request.form("p1_"&i)
				va_p2=request.form("p2_"&i)
				va_bon=request.form("bon_"&i)
				va_rec=request.form("rec_"&i)
		end if	
	else
			nu_matricula = request.form("nu_matricula_"&i)
			va_faltas=request.form("faltas_"&i)						
			va_t1=request.form("t1_"&i)
			va_t2=request.form("t2_"&i)
			va_p1=request.form("p1_"&i)
			va_p2=request.form("p2_"&i)
			va_bon=request.form("bon_"&i)
			va_rec=request.form("rec_"&i)
	end if
	
if fail = 0 then 		
Session("va_faltas")=va_faltas
Session("va_pt")=va_pt
Session("va_pp")=va_pp
Session("va_t1")=va_t1
Session("va_t2")=va_t2
Session("va_p1")=va_p1
Session("va_p2")=va_p2
'Session("va_bon")=va_bon
Session("va_bon")=""
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
								erro = "a1"
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
'				if va_pp1="" or isnull(va_pp1) then
'					if  fail = 1 then
'						grava = "no"
'					else
'							fail = 1 
'							erro = "pp1a"
'							matric_Erro=i
'							url = nu_matricula&"_"&va_p1&"_"&erro
'							grava = "no"
'					end if
'				else	
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
'				end if																				
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
'				if va_pp2="" or isnull(va_pp2) then
'					if  fail = 1 then
'						grava = "no"
'					else
'							fail = 1 
'							erro = "pp2a"
'							matric_Erro=i
'							url = nu_matricula&"_"&va_p2&"_"&erro
'							grava = "no"
'					end if
'				else
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
							url = nu_matricula&"_"&va_pp2&"_"&erro
							grava = "no"
						end if					
					end if
'				end if																				
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
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'PESOS
	if va_t1=NULL or va_t2=NULL then
		va_pt=NULL	
		s_va_pt=0
		soma_va_pt=0
	else		
		if va_pt="" or isnull(va_pt) then
			va_pt=1		
			s_va_pt=1
			soma_va_pt=1		
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
	end if	
		
	
	if (va_pp1="" or isnull(va_pp1)) and s_va_p1<>1 then
		va_pp1=NULL
		s_va_pp1=0
		soma_va_pp1=0
	elseif (va_pp1="" or isnull(va_pp1)) and s_va_p1=1 then		
		va_pp1=1
		s_va_pp1=1
		soma_va_pp1=1		
	else
		teste_va_pp1 = isnumeric(va_pp1)
		if teste_va_pp1= true then					
		va_pp1=va_pp1*1			
			if va_pp1 =<100 then
				IF Int(va_pp1)=va_pp1 THEN
					s_va_pp1=1
					soma_va_pp1=va_pp1						
				ELSE	
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "pp1"
						matric_Erro=i
						url = nu_matricula&"_"&va_pp1&"_"&erro
						grava = "no"
					end if					
				end if																				
			else
				if  fail = 1 then
					grava = "no"
				else
						fail = 1 
						erro = "pp1"
						matric_Erro=i
						url = nu_matricula&"_"&va_pp1&"_"&erro
						grava = "no"
				end if					
			end if				
		else
			if  fail = 1 then
				grava = "no"
			else
					fail = 1 
					erro = "pp1"
					matric_Erro=i
					url = nu_matricula&"_"&va_pp1&"_"&erro
					grava = "no"
			end if
		end if
	end if


	if (va_pp2="" or isnull(va_pp2)) and s_va_p2<>1  then
		va_pp2=NULL		
		s_va_pp2=0
		soma_va_pp2=0
	elseif (va_pp2="" or isnull(va_pp2)) and s_va_p2=1  then
		va_pp2=1		
		s_va_pp2=1
		soma_va_pp2=1			
	else
		teste_va_pp2 = isnumeric(va_pp2)
		if teste_va_pp2= true then					
			va_pp2=va_pp2*1			
			if va_pp2 =<100 then
				IF Int(va_pp2)=va_pp2 THEN
					s_va_pp2=1
					soma_va_pp2=va_pp2						
				ELSE	
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "pp2"
						matric_Erro=i
						url = nu_matricula&"_"&va_pp2&"_"&erro
						grava = "no"
					end if					
				end if																				
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "pp2"
					matric_Erro=i
					url = nu_matricula&"_"&va_pp2&"_"&erro
					grava = "no"
				end if					
			end if				
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "pp2"
				matric_Erro=i
				url = nu_matricula&"_"&va_pp2&"_"&erro
				grava = "no"
			end if
		end if
	end if
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
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
	soma_va_t1=soma_va_t1*1
	soma_va_t2=soma_va_t2*1
	
	divisor_mt=s_va_t1+s_va_t2
	dividendo_mt=soma_va_t1+soma_va_t2
	
		if divisor_mt=0 THEN
			media_t="no"
			mt=NULL						
		else
			media_t="ok"
				mt=dividendo_mt/divisor_mt
				mt_m1=mt
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

	if media_t="ok" and (s_va_p1=1 or s_va_p2=1) then
		mt_m1=mt_m1*1
		m1=((mt_m1*soma_va_pt)+(soma_va_p1*soma_va_pp1)+(soma_va_p2*soma_va_pp2))/(soma_va_pt+soma_va_pp1+soma_va_pp2)
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
				if  fail = 1 then
					grava = "no"
				else											
					fail = 1 
					erro = "m2"
					url = nu_matricula&"_"&va_bon&"_"&erro
					grava = "no"
				end if
			end if
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
		
		m3=m2
	end if
	
if grava = "ok" then
	
		'	response.Write("Select * from TB_Nota_B WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo)

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		CONEXAO0 = "Select * from TB_Nota_B WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
		Set RS0 = CON.Execute(CONEXAO0)
	If RS0.EOF THEN	
	
		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Nota_B", CON, 2, 2 'which table do you want open
		RS.addnew
		
			RS("CO_Matricula") = nu_matricula
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Periodo") = periodo
			RS("NU_Faltas")=va_faltas				
			RS("VA_Teste1")=va_t1
			RS("VA_Teste2")=va_t2
			RS("MD_Teste")=mt
			RS("PE_Teste")=va_pt
			RS("VA_Prova1")=va_p1
			RS("VA_Prova2")=va_p2
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
		CONEXAO0 = "DELETE * from TB_Nota_B WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
		Set RS0 = CON.Execute(CONEXAO0)

		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Nota_B", CON, 2, 2 'which table do you want open
		RS.addnew
		
			RS("CO_Matricula") = nu_matricula
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Periodo") = periodo
			RS("NU_Faltas")=va_faltas				
			RS("VA_Teste1")=va_t1
			RS("VA_Teste2")=va_t2
			RS("MD_Teste")=mt
			RS("PE_Teste")=va_pt
			RS("VA_Prova1")=va_p1
			RS("VA_Prova2")=va_p2
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

next
if fail = 1 then

response.Redirect(endereco_origem&"notas.asp?or=01&opt=err6&hp=err_"&url&"&obr="&obr) 

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