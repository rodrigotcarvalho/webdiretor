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
		ABRIR = "DBQ="& CAMINHO_nc & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CONMT = Server.CreateObject("ADODB.Connection") 
		ABRIRMT = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONMT.Open ABRIRMT
		
				
		Set RSMT  = Server.CreateObject("ADODB.Recordset")
		SQL_MT  = "Select CO_Materia_Principal from TB_Materia WHERE CO_Materia = '"& co_materia &"'"
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
	va_pt1=request.form("va_pt1")
	va_pt2=request.form("va_pt2")
	va_pt3=request.form("va_pt3")
	va_pt4=request.form("va_pt4")
	
	if nu_matricula = "falta" then
			i=i*1
			max=max*1
		if i=max then
		grava = "no"
		else
			while nu_matricula = "falta"
			i=i+1
			nu_matricula = request.form("nu_matricula_"&i)		
			va_t1=request.form("t1_"&i)
			va_t2=request.form("t2_"&i)
			va_t3=request.form("t3_"&i)
			va_t4=request.form("t4_"&i)						
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
			va_t4=request.form("t4_"&i)						
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
Session("t4")=va_t4
Session("p1")=va_p1
Session("p2")=va_p2
Session("bon")=va_bon
Session("rec")=va_rec	
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

'PESOS
	if va_pt1="" or isnull(va_pt1) then
		va_pt1=NULL		
		s_va_pt1=0
		soma_va_pt1=0		
	else
		teste_va_pt1 = isnumeric(va_pt1)
		if teste_va_pt1= true then					
			va_pt1=va_pt1*1			
			if va_pt1 =<10 then
				s_va_pt1=1
				soma_va_pt1=va_pt1																								
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "pt1"
					matric_Erro=i
					url = nu_matricula&"_"&va_pt1&"_"&erro
					grava = "no"
				end if					
			end if				
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "pt1"
				matric_Erro=i
				url = nu_matricula&"_"&va_pt1&"_"&erro
				grava = "no"
			end if
		end if
	end if
	
	if va_pt2="" or isnull(va_pt2) then
		va_pt2=NULL		
		s_va_pt2=0
		soma_va_pt2=0		
	else
		teste_va_pt2 = isnumeric(va_pt2)
		if teste_va_pt2= true then					
			va_pt2=va_pt2*1			
			if va_pt2 =<10 then
				s_va_pt2=1
				soma_va_pt2=va_pt2						
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "pt2"
					matric_Erro=i
					url = nu_matricula&"_"&va_pt2&"_"&erro
					grava = "no"
				end if					
			end if				
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "pt2"
				matric_Erro=i
				url = nu_matricula&"_"&va_pt2&"_"&erro
				grava = "no"
			end if
		end if
	end if	

	if va_pt3="" or isnull(va_pt3) then
		va_pt3=NULL		
		s_va_pt3=0
		soma_va_pt3=0		
	else
		teste_va_pt3 = isnumeric(va_pt3)
		if teste_va_pt3= true then					
			va_pt3=va_pt3*1			
			if va_pt3 =<10 then
				s_va_pt3=1
				soma_va_pt3=va_pt3						
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "pt3"
					matric_Erro=i
					url = nu_matricula&"_"&va_pt3&"_"&erro
					grava = "no"
				end if					
			end if				
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "pt3"
				matric_Erro=i
				url = nu_matricula&"_"&va_pt3&"_"&erro
				grava = "no"
			end if
		end if
	end if	

	if va_pt4="" or isnull(va_pt4) then
		va_pt4=NULL		
		s_va_pt4=0
		soma_va_pt4=0		
	else
		teste_va_pt4 = isnumeric(va_pt4)
		if teste_va_pt4= true then					
			va_pt4=va_pt4*1			
			if va_pt4 =<10 then
				s_va_pt4=1
				soma_va_pt4=va_pt4						
			else
				if  fail = 1 then
					grava = "no"
				else
						fail = 1 
						erro = "pt4"
						matric_Erro=i
						url = nu_matricula&"_"&va_pt4&"_"&erro
						grava = "no"
				end if					
			end if				
		else
			if  fail = 1 then
				grava = "no"
			else
					fail = 1 
					erro = "pt4"
					matric_Erro=i
					url = nu_matricula&"_"&va_pt4&"_"&erro
					grava = "no"
			end if
		end if
	end if	
soma_va_pt1=soma_va_pt1*1
soma_va_pt2=soma_va_pt2*1
soma_va_pt3=soma_va_pt3*1
soma_va_pt4=soma_va_pt4*1
total_peso=soma_va_pt1+soma_va_pt2+soma_va_pt3+soma_va_pt4

if total_peso>10 then
		fail = 1 
		erro = "pt"
		matric_Erro=i
		url = nu_matricula&"_"&total_peso&"_"&erro
		grava = "no"
end if		
	
	
	
'TESTES
	if va_t1="" or isnull(va_t1) then
		va_t1=NULL		
		s_va_t1=0
		soma_va_t1=0		
	else
		teste_va_t1 = isnumeric(va_t1)
		if teste_va_t1= true then					
			va_t1=va_t1*1			
			soma_va_pt1=soma_va_pt1*1
				if va_t1 =<soma_va_pt1 then
					s_va_t1=1
					soma_va_t1=va_t1																										
				else
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "t1-c$1"
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
					erro = "t1-c$0"
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
			soma_va_pt2=soma_va_pt2*1			
			if va_t2 =<soma_va_pt2 then
					s_va_t2=1
					soma_va_t2=va_t2																										
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "t2-c$1"
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
				erro = "t2-c$0"
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
			soma_va_pt3=soma_va_pt3*1	
			if va_t3 =<soma_va_pt3 then
					s_va_t3=1
					soma_va_t3=va_t3																										
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "t3-c$1"
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
				erro = "t3-c$0"
				matric_Erro=i
				url = nu_matricula&"_"&va_t3&"_"&erro
				grava = "no"
			end if
		end if
	end if
	if va_t4="" or isnull(va_t4) then
		va_t4=NULL		
		s_va_t4=0
		soma_va_t4=0		
	else
		teste_va_t4 = isnumeric(va_t4)
		if teste_va_t4= true then					
			va_t4=va_t4*1	
			soma_va_pt4=soma_va_pt4*1		
			if va_t4 =<soma_va_pt4 then
					s_va_t4=1
					soma_va_t4=va_t4																										
			else
				if  fail = 1 then
					grava = "no"
				else
						fail = 1 
						erro = "t4-c$1"
						matric_Erro=i
						url = nu_matricula&"_"&va_t4&"_"&erro
						grava = "no"
				end if					
			end if				
		else
			if  fail = 1 then
				grava = "no"
			else
					fail = 1 
					erro = "t4-c$0"
					matric_Erro=i
					url = nu_matricula&"_"&va_t4&"_"&erro
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
			if va_p1 =<10 then
					s_va_p1=1
					soma_va_p1=va_p1																									
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "p1-c$1"
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
					erro = "p1-c$0"
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
			if va_p2 =<10 then
					s_va_p2=1
					soma_va_p2=va_p2																			
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "p2-c$1"
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
				erro = "p2-c$0"
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
	
	if va_rec="" or isnull(va_rec) then
		va_rec=NULL		
		s_va_rec=0
		soma_va_rec=0		
	else
		teste_va_rec = isnumeric(va_rec) 
		if teste_va_rec = true then
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
	if periodo=4 then
		str=NULL
		if va_p2="" or isnull(va_p2) then
			m1=NULL
			m2=NULL
			m3=NULL	
		else	
			m1=soma_va_p2
			m2=soma_va_p2
			m3=soma_va_p2
				decimo = m3 - Int(m3)
				If decimo >= 0.75 Then
					nota_arredondada = Int(m3) + 1
					m3_arred=nota_arredondada
				elseIf decimo >= 0.25 Then
					nota_arredondada = Int(m3) + 0.5
					m3_arred=nota_arredondada
				else
					nota_arredondada = Int(m3)
					m3_arred=nota_arredondada											
				End If			
				m3 = formatNumber(m3_arred,1)	
		end if			
	else
		soma_va_t1=soma_va_t1*1
		soma_va_t2=soma_va_t2*1
		soma_va_t3=soma_va_t3*1
		soma_va_t4=soma_va_t4*1	
		soma_va_p1=soma_va_p1*1
		str=soma_va_t1+soma_va_t2+soma_va_t3+soma_va_t4
		
		
	
		if s_va_t1=0 and s_va_t2=0 and s_va_t3=0 and s_va_t4=0 and str=0 then
			str=NULL
			media_str="no"			
		else	
				if str >10 then
					if  fail = 1 then
						grava = "no"
					else
							fail = 1 
							erro = "str"
							matric_Erro=i
							url = nu_matricula&"_"&str&"_"&erro
							grava = "no"
					end if
					media_str="no"
				else
					media_str="ok"						
				end if
		end if		
	
		if (curso=1 and etapa=1 and s_va_p2=1) then
			m1_m3=soma_va_p2							
			m1 = formatNumber(m1_m3,1)		
		elseif (media_str="ok" and s_va_p1=1 and s_va_p2=1) then
			m1_m3=(str+soma_va_p1+soma_va_p2)/3							
			m1 = formatNumber(m1_m3,1)
		else
			m1=NULL
		end if	
			
		if isnull(m1) or m1="" then
			m2=NULL
			m3=NULL	
		else		
			if isnull(va_bon) or va_bon="" then
			m2=m1
			m2_m3=m1		
			else
				m1=m1*1		
				va_bon=va_bon*1
				m2_m3=m1_m3+va_bon
				
				if m2_m3>10 then
					if  fail = 1 then
						grava = "no"
					else											
						fail = 1 
						erro = "m2"
						url = nu_matricula&"_"&va_bon&"_"&erro
						grava = "no"
					end if
				end if				
					m2 = formatNumber(m2_m3,1)
			end if
	
			m3=m2_m3
	
				decimo = m3 - Int(m3)
				If decimo >= 0.75 Then
					nota_arredondada = Int(m3) + 1
					m3_arred=nota_arredondada
				elseIf decimo >= 0.25 Then
					nota_arredondada = Int(m3) + 0.5
					m3_arred=nota_arredondada
				else
					nota_arredondada = Int(m3)
					m3_arred=nota_arredondada											
				End If			
				m3 = formatNumber(m3_arred,1)
		end if
	end if	
'response.Write(i&"-"&va_bon&"<BR>")
if i=10 then
'response.end()
end if	
if grava = "ok" then
	
		'	response.Write("Select * from TB_Nota_C WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo)

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		CONEXAO0 = "Select * from TB_Nota_C WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
		Set RS0 = CON.Execute(CONEXAO0)
	If RS0.EOF THEN	
	
		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Nota_C", CON, 2, 2 'which table do you want open
		RS.addnew
		
			RS("CO_Matricula") = nu_matricula
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Periodo") = periodo
			RS("NU_Faltas")=va_faltas				
			RS("VA_Pt1")=va_pt1
			RS("VA_Pt2")=va_pt2
			RS("VA_Pt3")=va_pt3
			RS("VA_Pt4")=va_pt4
			RS("VA_Tr1")=va_t1
			RS("VA_Tr2")=va_t2
			RS("VA_Tr3")=va_t3
			RS("VA_Tr4")=va_t4
			RS("VA_Str")=str
			RS("VA_TE")=va_p1
			RS("VA_PF")=va_p2						
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
		CONEXAO0 = "DELETE * from TB_Nota_C WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
		Set RS0 = CON.Execute(CONEXAO0)

		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Nota_C", CON, 2, 2 'which table do you want open
		RS.addnew
		
			RS("CO_Matricula") = nu_matricula
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Periodo") = periodo
			RS("NU_Faltas")=va_faltas	
			RS("VA_Pt1")=va_pt1
			RS("VA_Pt2")=va_pt2
			RS("VA_Pt3")=va_pt3
			RS("VA_Pt4")=va_pt4						
			RS("VA_Tr1")=va_t1
			RS("VA_Tr2")=va_t2
			RS("VA_Tr3")=va_t3
			RS("VA_Tr4")=va_t4
			RS("VA_Str")=str
			RS("VA_TE")=va_p1
			RS("VA_PF")=va_p2						
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