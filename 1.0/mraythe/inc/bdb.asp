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
			va_faltas=request.form("faltas_"&i)					
			va_tr1=request.form("tr1_"&i)
			va_tr2=request.form("tr2_"&i)
			va_tr3=request.form("tr3_"&i)
			va_tr4=request.form("tr4_"&i)			
			va_t1=request.form("t1_"&i)
			va_t2=request.form("t2_"&i)
			va_t3=request.form("t3_"&i)
			va_t4=request.form("t4_"&i)						
			va_p1=request.form("p1_"&i)
			va_bon=request.form("bon_"&i)
			va_rec=request.form("rec_"&i)
			wend
		end if	
	else
			nu_matricula = request.form("nu_matricula_"&i)
			va_faltas=request.form("faltas_"&i)						
			va_tr1=request.form("tr1_"&i)
			va_tr2=request.form("tr2_"&i)
			va_tr3=request.form("tr3_"&i)
			va_tr4=request.form("tr4_"&i)			
			va_t1=request.form("t1_"&i)
			va_t2=request.form("t2_"&i)
			va_t3=request.form("t3_"&i)
			va_t4=request.form("t4_"&i)						
			va_p1=request.form("p1_"&i)
			va_bon=request.form("bon_"&i)
			va_rec=request.form("rec_"&i)
	end if
	
if fail = 0 then 		
	Session("faltas")=va_faltas
	Session("tr1")=va_tr1
	Session("tr2")=va_tr2
	Session("tr3")=va_tr3
	Session("tr4")=va_tr4
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
s_va_t1=0

	if va_tr1="" or isnull(va_tr1) then
		va_tr1=NULL		
		s_va_tr1=0
		soma_va_tr1=0
	else
		teste_va_tr1 = isnumeric(va_tr1)
		if teste_va_tr1= true then	
			soma_va_pt1=soma_va_pt1*1	
			va_tr1=va_tr1*1						
			if va_tr1 =< soma_va_pt1 then
				soma_va_tr1=va_tr1	
				s_va_tr1=1											
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "tr1-b$1"
					matric_Erro=i
					url = nu_matricula&"_"&va_tr1&"_"&erro
					grava = "no"
				end if	
			end if						
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "tr1-b$0"
				matric_Erro=i
				url = nu_matricula&"_"&va_tr1&"_"&erro
				grava = "no"
			end if					
		end if			
	end if

	if va_tr2="" or isnull(va_tr2) then
		va_tr2=NULL		
		s_va_tr2=0
		soma_va_tr2=0
	else			
		teste_va_tr2 = isnumeric(va_tr2)
		if teste_va_tr2= true then		
			soma_va_pt2=soma_va_pt2*1	
			va_tr2=va_tr2*1					
			if va_tr2 =< soma_va_pt2 then
				soma_va_tr2=va_tr2										
				s_va_tr2=1
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "tr2-b$1"
					matric_Erro=i
					url = nu_matricula&"_"&va_tr1&"_"&erro
					grava = "no"
				end if	
			end if			
		else
			if  fail = 1 then
				grava = "no"
			else
					fail = 1 
					erro = "tr2-b$0"
					matric_Erro=i
					url = nu_matricula&"_"&va_tr2&"_"&erro
					grava = "no"
			end if		
		end if
	end if
	
	if va_tr3="" or isnull(va_tr3) then
		va_tr3=NULL		
		s_va_tr3=0
		soma_va_tr3=0
	else			
		teste_va_tr3 = isnumeric(va_tr3)
		if teste_va_tr3= true then		
			soma_va_pt3=soma_va_pt3*1	
			va_tr3=va_tr3*1				
			if va_tr3 =< soma_va_pt3 then
				soma_va_tr3=va_tr3									
				s_va_tr3=1
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "tr3-b$1"
					matric_Erro=i
					url = nu_matricula&"_"&va_tr1&"_"&erro
					grava = "no"
				end if	
			end if				
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "tr3-b$0"
				matric_Erro=i
				url = nu_matricula&"_"&va_tr3&"_"&erro
				grava = "no"
			end if		
		end if
	end if	
	
	if va_tr4="" or isnull(va_tr4) then
		va_tr4=NULL		
		s_va_tr4=0
		soma_va_tr4=0
	else			
		teste_va_tr4 = isnumeric(va_tr4)
		if teste_va_tr4= true then		
			soma_va_pt4=soma_va_pt4*1	
			va_tr4=va_tr4*1				
			if va_tr4 =< soma_va_pt4 then
				soma_va_tr4=va_tr4								
				s_va_tr4=1
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "tr4-b$1"
					matric_Erro=i
					url = nu_matricula&"_"&va_tr1&"_"&erro
					grava = "no"
				end if	
			end if			
		else
			if  fail = 1 then
				grava = "no"
			else
					fail = 1 
					erro = "tr4-b$0"
					matric_Erro=i
					url = nu_matricula&"_"&va_tr4&"_"&erro
					grava = "no"
			end if		
		end if
	end if
			
	if va_t1="" or isnull(va_t1) then
		va_t1=NULL		
		s_va_t1=0
		soma_va_t1=0		
	else
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
						erro = "t1-b$1"
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
					erro = "t1-b$0"
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
			if va_t2 =<10 then
					s_va_t2=1
					soma_va_t2=va_t2																										
			else
				if  fail = 1 then
					grava = "no"
				else
						fail = 1 
						erro = "t2-b$1"
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
				erro = "t2-b$01"
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
			if va_t3 =<10 then
					s_va_t3=1
					soma_va_t3=va_t3																										
			else
				if  fail = 1 then
					grava = "no"
				else
						fail = 1 
						erro = "t2-b$1"
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
					erro = "t3-b$0"
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
			if va_t4 =<10 then
					s_va_t4=1
					soma_va_t4=va_t4																										
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "t4-b$1"
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
				erro = "t4-b$0"
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
					erro = "p1-b$1"
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
				erro = "p1-b$0"
				matric_Erro=i
				url = nu_matricula&"_"&va_p1&"_"&erro
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
		soma_va_tr1=soma_va_tr1*1
		soma_va_tr2=soma_va_tr2*1
		soma_va_tr3=soma_va_tr3*1
		soma_va_tr4=soma_va_tr4*1	
		soma_va_p1=soma_va_p1*1
		
		str=soma_va_tr1+soma_va_tr2+soma_va_tr3+soma_va_tr4
	
		if s_va_tr1=0 and s_va_tr2=0 and s_va_tr3=0 and s_va_tr4=0 then
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
	'response.Write(str&"<BR>")				
	'response.End()
		s_va_t1=s_va_t1*1
		s_va_t2=s_va_t2*1
		s_va_t3=s_va_t3*1
		s_va_t4=s_va_t4*1	
		soma_va_t1=soma_va_t1*1
		soma_va_t2=soma_va_t2*1
		soma_va_t3=soma_va_t3*1
		soma_va_t4=soma_va_t4*1
			
		divisor_mt=s_va_t1+s_va_t2+s_va_t3+s_va_t4
		dividendo_mt=soma_va_t1+soma_va_t2+soma_va_t3+soma_va_t4
		
			if divisor_mt=0 THEN
				media_t="no"
				mt=NULL						
			else
				media_t="ok"
					mt=dividendo_mt/divisor_mt				
					mt = formatNumber(mt,1)					
			end if
	
		s_va_p1=s_va_p1*1
		str=str*1
		mt=mt*1	
		soma_va_p1=soma_va_p1*1	
		if (co_materia="ART" or co_materia="EDUF") and  media_str="ok" then
			m1_m3=str							
			m1 = formatNumber(m1_m3,1)		
		elseif (co_materia="EDUR") and  media_str="ok" and s_va_p1=1 then
			m1_m3=(str+soma_va_p1)/2
								
			m1 = formatNumber(m1_m3,1)	
		elseif media_str="ok" and media_t="ok" and s_va_p1=1 then
			m1_m3=(str+mt+soma_va_p1)/3							
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
					m2 = formatNumber(m2,1)
			end if
		if va_rec="" or isnull(va_rec) then		
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
	
			if co_materia= "EDUR" and  media_str="ok" and s_va_p1=1 then
			'if i=5 then
			'response.Write(str&"-"&soma_va_rec&"-"&soma_va_p1)
			'response.end()
			'end if
	
	
					if soma_va_rec<=str  and soma_va_rec<=soma_va_p1 then							
						m3=m2
					elseif soma_va_p1>str and soma_va_rec>str then
						m3=((soma_va_rec+soma_va_p1)/2)+soma_va_bon
						'response.Write(m3&"=(("&soma_va_rec&"+"&soma_va_p1&")/2)+"&soma_va_bon)
						'response.end()
					elseif soma_va_p1<str and soma_va_rec>soma_va_p1 then
						m3=((str+soma_va_rec)/2)+soma_va_bon
					elseif soma_va_p1=str and soma_va_rec>soma_va_p1 then
						m3=((str+soma_va_rec)/2)+soma_va_bon
					end if		
			else
				if (co_materia="ART" or co_materia="EDUF") and  media_str="ok" then
					if str>soma_va_rec then							
						m3=m2
					else
						m3=soma_va_rec
					end if	
				else
					if media_t="ok" and s_va_p1=1 then
					mt=mt*1	
					soma_va_p1=soma_va_p1*1	
						if (mt>soma_va_p1) and (soma_va_rec>soma_va_p1) then
							m3=((str+mt+soma_va_rec)/3)+soma_va_bon
							'response.Write(i&"b<BR>")
						elseif (soma_va_p1>mt) and (soma_va_rec>mt) then
							m3=((str+soma_va_rec+soma_va_p1)/3)+soma_va_bon
							'response.Write(i&"c<BR>")
						elseif (mt=soma_va_p1) and (soma_va_rec>mt) then
							m3=((str+soma_va_rec+soma_va_p1)/3)+soma_va_bon
							'response.Write(i&"c<BR>")						
						else
							m3=m2
						end if
					else
					m3=m2
					'response.Write(i&"d<BR>")
					end if
				end if
			end if		
		end if
'			decimo = m3 - Int(m3)
'			If decimo >= 0.75 Then
'				nota_arredondada = Int(m3) + 1
'				m3_arred=nota_arredondada
'			elseIf decimo >= 0.25 Then
'				nota_arredondada = Int(m3) + 0.5
'				m3_arred=nota_arredondada
'			else
'				nota_arredondada = Int(m3)
'				m3_arred=nota_arredondada											
'			End If			
'			m3 = formatNumber(m3_arred,1)
			m3=arredonda(m3,"mat_dez",1,outro)	

	end if
end if
				
if isnumeric(m3) and m3>10 then
	m3=10
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
			RS("VA_Pt1")=va_pt1
			RS("VA_Pt2")=va_pt2
			RS("VA_Pt3")=va_pt3
			RS("VA_Pt4")=va_pt4						
			RS("VA_Tr1")=va_tr1
			RS("VA_Tr2")=va_tr2
			RS("VA_Tr3")=va_tr3
			RS("VA_Tr4")=va_tr4			
			RS("VA_Str")=str
			RS("VA_Te1")=va_t1
			RS("VA_Te2")=va_t2
			RS("VA_Te3")=va_t3
			RS("VA_Te4")=va_t4			
			RS("VA_Mte")=mt
			RS("VA_Pr")=va_p1									
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
			RS("VA_Pt1")=va_pt1
			RS("VA_Pt2")=va_pt2
			RS("VA_Pt3")=va_pt3
			RS("VA_Pt4")=va_pt4						
			RS("VA_Tr1")=va_tr1
			RS("VA_Tr2")=va_tr2
			RS("VA_Tr3")=va_tr3
			RS("VA_Tr4")=va_tr4			
			RS("VA_Str")=str
			RS("VA_Te1")=va_t1
			RS("VA_Te2")=va_t2
			RS("VA_Te3")=va_t3
			RS("VA_Te4")=va_t4			
			RS("VA_Mte")=mt
			RS("VA_Pr")=va_p1									
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