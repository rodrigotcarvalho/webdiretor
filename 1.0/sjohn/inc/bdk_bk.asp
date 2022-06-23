<%'On Error Resume Next%>
<!--#include file="caminhos.asp"-->
<!--#include file="funcoes.asp"--> 
<!--#include file="bd_modelo_k.asp"-->
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
ABRIR = "DBQ="& CAMINHO_nk & ";Driver={Microsoft Access Driver (*.mdb)}"
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
			wend
			va_av1=request.form("av1_"&i)
			va_av2=request.form("av2_"&i)
			va_av3=request.form("av3_"&i)
			va_av4=request.form("av4_"&i)
			va_av5=request.form("av5_"&i)
			va_sim=request.form("sim_"&i)
			va_rs=request.form("rs_"&i)
			va_bat=request.form("bat_"&i)				
			va_rb=request.form("rb_"&i)					
			va_bon=request.form("bon_"&i)
			va_rec=request.form("rec_"&i)
		end if	
	
	else
		nu_matricula = request.form("nu_matricula_"&i)
		va_av1=request.form("av1_"&i)
		va_av2=request.form("av2_"&i)
		va_av3=request.form("av3_"&i)
		va_av4=request.form("av4_"&i)
		va_av5=request.form("av5_"&i)
		va_sim=request.form("sim_"&i)
		va_rs=request.form("rs_"&i)
		va_bat=request.form("bat_"&i)				
		va_rb=request.form("rb_"&i)					
		va_bon=request.form("bon_"&i)
		va_rec=request.form("rec_"&i)
	end if
	
	if fail = 0 then 		
		Session("va_av1")=va_av1
		Session("va_av2")=va_av2
		Session("va_av3")=va_av3
		Session("va_av4")=va_av4
		Session("va_av5")=va_av5
		Session("va_sim")=va_sim
		Session("va_bat")=va_bat
		Session("va_bon")=va_bon
		Session("va_rec")=va_rec		
	end if
	
	if isnull(va_rs) or va_rs="" then
		va_rs="N"
	end if
	
	if isnull(va_rb) or va_rb="" then
		va_rb="N"
	end if	
	
'////////////////////////////////////////////////////////////////
'pesos (por enquanto essa verificação não é usada)
	
	if va_pt="" or isnull(va_pt) then
		va_pt = 1
		'p_va_pt="vazio"
		teste_va_pt= true
	else
		teste_va_pt = isnumeric(va_pt)
	end if
	
	if va_pp="" or isnull(va_pp) then
		if etapa <>"V" then
			etapa=etapa*1
				if etapa=3 then
					va_pp = 4
				else	
					va_pp = 2		
				end if
				'p_va_pp="vazio"
				teste_va_pp= true
		else
			va_pp = 2
			teste_va_pp= true	
		end if		
	else
		teste_va_pp = isnumeric(va_pp)
	end if
	
	
	if teste_va_pt=true and teste_va_pp=true then
		va_pt=va_pt*1
		va_pp=va_pp*1
	
	else
		fail = 1 
		erro = "pt"
		url = 0&"_"&sum_p&"_"&erro
		grava = "no"
	end if

'///////////////////////////////////////////////////////////////////////////

'TESTES
	s_va_t=0
	if va_av1="" or isnull(va_av1) then
		va_av1=NULL		
		s_va_av1=0
		soma_av1=0	
		av1_lancado="no"			
	else
		teste_va_av1 = isnumeric(va_av1)
		if teste_va_av1= true then					
			va_av1=va_av1*1			
			if va_av1 =<100 then
				IF Int(va_av1)=va_av1 THEN
					s_va_av1=1
					soma_av1=va_av1
					av1_lancado="sim"													
				ELSE	
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "av1"
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
					erro = "av1"
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
				erro = "av1"
				url = nu_matricula&"_"&va_av1&"_"&erro
				grava = "no"
			end if
		end if
	end if

	if va_av2="" or isnull(va_av2) then
		va_av2=NULL		
		s_va_av2=0
		soma_av2=0	
		av2_lancado="no"			
	else
		teste_va_av2 = isnumeric(va_av2)
		if teste_va_av2= true then					
			va_av2=va_av2*1			
			if va_av2 =<100 then
				IF Int(va_av2)=va_av2 THEN
					s_va_av2=1
					soma_av2=va_av2
					av2_lancado="sim"													
				ELSE	
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "av2"
						matric_Erro=i
						url = nu_matricula&"_"&va_av2&"_"&erro
						grava = "no"
					end if					
				end if																				
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "av2"
					matric_Erro=i
					url = nu_matricula&"_"&va_av2&"_"&erro
					grava = "no"
				end if					
			end if				
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "av2"
				url = nu_matricula&"_"&va_av2&"_"&erro
				grava = "no"
			end if
		end if
	end if
	
	if va_av3="" or isnull(va_av3) then
		va_av3=NULL		
		s_va_av3=0
		soma_av3=0	
		av3_lancado="no"			
	else
		teste_va_av3 = isnumeric(va_av3)
		if teste_va_av3= true then					
			va_av3=va_av3*1			
			if va_av3 =<100 then
				IF Int(va_av3)=va_av3 THEN
					s_va_av3=1
					soma_av3=va_av3
					av3_lancado="sim"													
				ELSE	
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "av3"
						matric_Erro=i
						url = nu_matricula&"_"&va_av3&"_"&erro
						grava = "no"
					end if					
				end if																				
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "av3"
					matric_Erro=i
					url = nu_matricula&"_"&va_av3&"_"&erro
					grava = "no"
				end if					
			end if				
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "av3"
				url = nu_matricula&"_"&va_av3&"_"&erro
				grava = "no"
			end if
		end if
	end if	

	if va_av4="" or isnull(va_av4) then
		va_av4=NULL		
		s_va_av4=0
		soma_av4=0	
		av4_lancado="no"			
	else
		teste_va_av4 = isnumeric(va_av4)
		if teste_va_av4= true then					
			va_av4=va_av4*1			
			if va_av4 =<100 then
				IF Int(va_av4)=va_av4 THEN
					s_va_av4=1
					soma_av4=va_av4
					av4_lancado="sim"													
				ELSE	
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "av4"
						matric_Erro=i
						url = nu_matricula&"_"&va_av4&"_"&erro
						grava = "no"
					end if					
				end if																				
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "av4"
					matric_Erro=i
					url = nu_matricula&"_"&va_av4&"_"&erro
					grava = "no"
				end if					
			end if				
		else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "av4"
					url = nu_matricula&"_"&va_av4&"_"&erro
					grava = "no"
				end if
		end if
	end if
	
	if va_av5="" or isnull(va_av5) then
		va_av5=NULL		
		s_va_av5=0
		soma_av5=0	
		av5_lancado="no"			
	else
		teste_va_av5 = isnumeric(va_av5)
		if teste_va_av5= true then					
			va_av5=va_av5*1			
			if va_av5 =<100 then
				IF Int(va_av5)=va_av5 THEN
					s_va_av5=1
					soma_av5=va_av5
					av5_lancado="sim"													
				ELSE	
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "av5"
						matric_Erro=i
						url = nu_matricula&"_"&va_av5&"_"&erro
						grava = "no"
					end if					
				end if																				
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "av5"
					matric_Erro=i
					url = nu_matricula&"_"&va_av5&"_"&erro
					grava = "no"
				end if					
			end if				
		else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "av5"
					url = nu_matricula&"_"&va_av5&"_"&erro
					grava = "no"
				end if
		end if
	end if
'response.Write(i&"-"&nu_matricula&"-"&va_apr7 &">"& va_v_apr7&"<BR>")

'//////////////////////////////////////////////////////////////////////
'Notas
	s_va_p=0
	if va_sim="" or isnull(va_sim) then
		va_sim=NULL
		s_va_sim=0		
		soma_sim=0	
		sim_lancado="no"	
	else
		teste_va_sim = isnumeric(va_sim)
		if teste_va_sim= true then					
			va_sim=va_sim*1			
			if va_sim =<10 then
				IF Int(va_sim)=va_sim THEN
					s_va_sim=1
					soma_sim=va_sim	
					sim_lancado="sim"					
				ELSE	
					if  fail = 1 then
						grava = "no"
					else					
						fail = 1 
						erro = "sim"
						url = nu_matricula&"_"&va_sim&"_"&erro
						grava = "no"
					end if					
				end if															
			else
				fail = 1 
				erro = "simv"
				url = nu_matricula&"_"&va_sim&"_"&erro
				grava = "no"
			end if				
		else
			fail = 1 
			erro = "sim"
			url = nu_matricula&"_"&va_sim&"_"&erro
			grava = "no"
		end if
	end if
	
	if va_bat="" or isnull(va_bat) then
		va_bat=NULL
		s_va_bat=0		
		soma_bat=0	
		bat_lancado="no"	
	else
		teste_va_bat = isnumeric(va_bat)
		if teste_va_bat= true then					
			va_bat=va_bat*1			
			if va_bat =<5 then
				IF Int(va_bat)=va_bat THEN
					s_va_bat=1
					soma_bat=va_bat	
					bat_lancado="sim"					
				ELSE	
					if  fail = 1 then
						grava = "no"
					else					
						fail = 1 
						erro = "bat"
						url = nu_matricula&"_"&va_bat&"_"&erro
						grava = "no"
					end if					
				end if															
			else
				fail = 1 
				erro = "batv"
				url = nu_matricula&"_"&va_bat&"_"&erro
				grava = "no"
			end if				
		else
			fail = 1 
			erro = "bat"
			url = nu_matricula&"_"&va_bat&"_"&erro
			grava = "no"
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
		periodo=periodo*1
		if periodo=1 or periodo=3 then 
	
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
	end if	

'/////////////////////////////////////////////////////////////////////////
'Médias


	if grava = "ok" then

		'if sim_lancado="sim" and va_rs="S" THEN
		if va_rs="S" THEN
			replica_sim=ReplicaInformacoes(unidade, curso, etapa, turma, nu_matricula, periodo, "TB_Nota_K", "VA_Sim", va_sim)
		end if
		
		'if bat_lancado="sim" and va_rb="S" THEN
		if va_rb="S" THEN		
			replica_sim=ReplicaInformacoes(unidade, curso, etapa, turma, nu_matricula, periodo, "TB_Nota_K", "VA_Bat", va_bat)
		end if	
	
	
		soma_av1=soma_av1*1
		soma_av2=soma_av2*1
		soma_av3=soma_av3*1
		soma_av4=soma_av4*1
		soma_av5=soma_av5*1
	
		s_va_av1=s_va_av1*1
		s_va_av2=s_va_av2*1
		s_va_av3=s_va_av3*1
		s_va_av4=s_va_av4*1
		s_va_av5=s_va_av5*1
		
		s_va_t=s_va_av1+s_va_av2+s_va_av3+s_va_av4+s_va_av5

'response.Write("if "&teste_1_lancado&"="no" and "&teste_2_lancado&"="no" and "&teste_3_lancado&"="no" and "&teste_4_lancado&"="no" then<BR>")
		'if ((av1_lancado="no" or av2_lancado="no") and  (etapa =1 or etapa =2)) or (av2_lancado="no" and etapa =3)   THEN
		if (av1_lancado="no" or av2_lancado="no") THEN
			media_av="no"
			mav=NULL							
		else
			media_av="ok"		
			mav=(soma_av1+soma_av2+soma_av3+soma_av4+soma_av5)/s_va_t
		
		'mt=mt*10
			decimo = mav - Int(mav)
			If decimo >= 0.5 Then
				nota_arredondada = Int(mav) + 1
				mav=nota_arredondada
			Else
				nota_arredondada = Int(mav)
				mav=nota_arredondada					
			End If
			'mt=mt/10			
			mav = formatNumber(mav,0)					
		end if
		
		
	
		soma_sim=soma_sim*1
		soma_bat=soma_bat*1
		s_va_sim=s_va_sim*1
		s_va_bat=s_va_bat*1
		
	
		if media_av="ok" then
			mav=mav*1
			soma_sim=soma_sim*1
			soma_bat=soma_bat*1		
			m1=mav+soma_sim+soma_bat
			if m1>100 then
				m1=100
			end if	
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
		else
			m1=NULL		
		END IF
	
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
				m2=m2*1
				va_rec=va_rec*1
				m3_temp=(m2+va_rec)/2
				
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
	

	Set RS0 = Server.CreateObject("ADODB.Recordset")
	CONEXAO0 = "Select * from TB_Nota_K WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
	response.Write(CONEXAO0&"<BR>")
	Set RS0 = CON.Execute(CONEXAO0)
	
	If RS0.EOF THEN	
			
		'response.Write("4"&turma &"/"&co_materia_pr)
		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Nota_K", CON, 2, 2 'which table do you want open
		RS.addnew
	response.Write(nu_matricula&"-"&co_materia_pr&"-"&co_materia&"-"&periodo&"<BR>")		
			RS("CO_Matricula") = nu_matricula
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Periodo") = periodo	
			RS("VA_Av1")=va_av1
			RS("VA_Av2")=va_av2
			RS("VA_Av3")=va_av3
			RS("VA_Av4")=va_av4
			RS("VA_Av5")=va_av5
			RS("VA_Mav")=mav
			RS("VA_Sim")=va_sim
			RS("VA_Bat")=va_bat		
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
		CONEXAO0 = "DELETE * from TB_Nota_K WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
		Set RS0 = CON.Execute(CONEXAO0)

		Set RS = server.createobject("adodb.recordset")		
		RS.open "TB_Nota_K", CON, 2, 2 'which table do you want open
		RS.addnew
		
			RS("CO_Matricula") = nu_matricula
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Periodo") = periodo	
			RS("VA_Av1")=va_av1
			RS("VA_Av2")=va_av2
			RS("VA_Av3")=va_av3
			RS("VA_Av4")=va_av4
			RS("VA_Av5")=va_av5
			RS("VA_Mav")=mav
			RS("VA_Sim")=va_sim
			RS("VA_Bat")=va_bat	
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
		
'		sql_atualiza= "UPDATE TB_Nota_A SET VA_Teste1 ="&sql_va_av1&", VA_Teste2 ="&sql_va_t2&", VA_Teste3 ="&sql_va_t3&", VA_Teste4 ="&sql_va_t4&", MD_Teste =FORMAT("&sql_mt&",2), "
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
response.Redirect(endereco_origem&"comunicar.asp?or=01&nota=TB_Nota_K&opt=ok&obr="&obr)
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
