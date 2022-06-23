<%On Error Resume Next%>
<!--#include file="caminhos.asp"-->
<!--#include file="parametros.asp"-->
<!--#include file="funcoes.asp"-->
<%
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
		SQL_MT  = "Select CO_Materia_Principal from TB_Materia WHERE CO_Materia = '"& co_materia&"'"
		Set RSMT  = CONMT.Execute(SQL_MT)
		
co_materia_pr = RSMT("CO_Materia_Principal")
		
if Isnull(co_materia_pr) then
co_materia_pr= co_materia
else
co_materia_pr = co_materia_pr
end if

'////////////////////////////////////////////////////
'Nomes
dados_tabela=dados_planilha_notas(ano_letivo,unidade,curso,etapa,turma,co_materia,0,periodo,"C",outro)

	dados_separados=split(dados_tabela,"#$#")
	ln_pesos_vars=dados_separados(2)
	nm_pesos_vars=dados_separados(3)
	nm_vars=dados_separados(5)
	nm_bd=dados_separados(6)

	linha_pesos_variaveis=split(ln_pesos_vars,"#!#")
	nome_pesos_variaveis=split(nm_pesos_vars,"#!#")
	nome_variaveis=split(nm_vars,"#!#")
	variaveis_bd=split(nm_bd,"#!#")

'Nomes

no_v_apr1=nome_pesos_variaveis(2)
no_v_apr2=nome_pesos_variaveis(3)
no_v_apr3=nome_pesos_variaveis(4)
no_v_apr4=nome_pesos_variaveis(5)
no_v_apr5=nome_pesos_variaveis(6)
no_v_apr6=nome_pesos_variaveis(7)
no_v_apr7=nome_pesos_variaveis(8)
no_v_apr8=nome_pesos_variaveis(9)

no_apr1=nome_variaveis(0)
no_apr2=nome_variaveis(1)
no_apr3=nome_variaveis(2)
no_apr4=nome_variaveis(3)
no_apr5=nome_variaveis(4)
no_apr6=nome_variaveis(5)
no_apr7=nome_variaveis(6)
no_apr8=nome_variaveis(7)
no_sapr=nome_variaveis(8)
no_pr=nome_variaveis(9)
if periodo=4 then
	no_me=nome_variaveis(10)
	no_mc=nome_variaveis(11)
else	
	no_te=nome_variaveis(10)
	no_bon=nome_variaveis(11)
	no_me=nome_variaveis(12)
	no_mc=nome_variaveis(13)
	no_faltas=nome_variaveis(14)
end if	

fail = 0
for i=1 to max
grava="ok"
va_me=0
va_mc=0


nu_matricula = request.form("nu_matricula_"&i)



if nu_matricula = "falta" then
i=i+1
nu_matricula = request.form("nu_matricula_"&i)
va_apr1=request.form(no_apr1&"_"&i)
va_apr2=request.form(no_apr2&"_"&i)
va_apr3=request.form(no_apr3&"_"&i)
va_apr4=request.form(no_apr4&"_"&i)
va_apr5=request.form(no_apr5&"_"&i)
va_apr6=request.form(no_apr6&"_"&i)
va_apr7=request.form(no_apr7&"_"&i)
va_apr8=request.form(no_apr8&"_"&i)
va_v_apr1=request.form(no_v_apr1)
va_v_apr2=request.form(no_v_apr2)
va_v_apr3=request.form(no_v_apr3)
va_v_apr4=request.form(no_v_apr4)
va_v_apr5=request.form(no_v_apr5)
va_v_apr6=request.form(no_v_apr6)
va_v_apr7=request.form(no_v_apr7)
va_v_apr8=request.form(no_v_apr8)
'va_sapr=request.form(no_sapr&"_"&i)
va_pr=request.form(no_pr&"_"&i)
va_te=request.form(no_te&"_"&i)
va_bon=request.form(no_bon&"_"&i)
'va_me=request.form(no_me&"_"&i)
'va_mc=request.form(no_mc&"_"&i)
faltas=request.form(no_faltas&"_"&i)
else



va_apr1=request.form(no_apr1&"_"&i)
va_apr2=request.form(no_apr2&"_"&i)
va_apr3=request.form(no_apr3&"_"&i)
va_apr4=request.form(no_apr4&"_"&i)
va_apr5=request.form(no_apr5&"_"&i)
va_apr6=request.form(no_apr6&"_"&i)
va_apr7=request.form(no_apr7&"_"&i)
va_apr8=request.form(no_apr8&"_"&i)
va_v_apr1=request.form(no_v_apr1)
va_v_apr2=request.form(no_v_apr2)
va_v_apr3=request.form(no_v_apr3)
va_v_apr4=request.form(no_v_apr4)
va_v_apr5=request.form(no_v_apr5)
va_v_apr6=request.form(no_v_apr6)
va_v_apr7=request.form(no_v_apr7)
va_v_apr8=request.form(no_v_apr8)
'va_sapr=request.form(no_sapr&"_"&i)
va_pr=request.form(no_pr&"_"&i)
va_te=request.form(no_te&"_"&i)
va_bon=request.form(no_bon&"_"&i)
'va_me=request.form(no_me&"_"&i)
'va_mc=request.form(no_mc&"_"&i)
faltas=request.form(no_faltas&"_"&i)

end if
va_sapr=""
va_me=""
va_mc=""
		'teste_va_sapr= isnumeric(va_sapr)
		'teste_va_me = isnumeric(va_me)
		'teste_va_mc= isnumeric(va_mc)


	if faltas="" or isnull(faltas) then
		teste_faltas = true
		faltas=0
	else
		teste_faltas = isnumeric(faltas)
	end if	
	
	
	if teste_faltas= true then
			
	else
		fail = 1 
		erro = "f"
		url = nu_matricula&"_"&faltas&"_"&erro
		grava = "no"
	end if	
		
'////////////////////////////////////////////////////////////////
'pesos

	if va_v_apr1="" or isnull(va_v_apr1) then
		va_v_apr1 = 0
		p_v_apr1="vazio"
		teste_va_v_apr1= true
	else
		teste_va_v_apr1 = isnumeric(va_v_apr1)
	end if

	if va_v_apr2="" or isnull(va_v_apr2) then
		va_v_apr2 = 0
		p_v_apr2="vazio"
		teste_va_v_apr2= true
	else
		teste_va_v_apr2= isnumeric(va_v_apr2)
	end if

	if va_v_apr3="" or isnull(va_v_apr3) then
		va_v_apr3 = 0
		p_v_apr3="vazio"
		teste_va_v_apr3= true
	else
		teste_va_v_apr3 = isnumeric(va_v_apr3)
	end if

	if va_v_apr4="" or isnull(va_v_apr4) then
		va_v_apr4 = 0
		p_v_apr4="vazio"
		teste_va_v_apr4= true
	else
		teste_va_v_apr4 = isnumeric(va_v_apr4)
	end if

	if va_v_apr5="" or isnull(va_v_apr5) then
		va_v_apr5 = 0
		p_v_apr5="vazio"
		teste_va_v_apr5= true
	else
		teste_va_v_apr5 = isnumeric(va_v_apr5)		
	end if

	if va_v_apr6="" or isnull(va_v_apr6) then
		va_v_apr6 = 0
		p_v_apr6="vazio"
		teste_va_v_apr6= true
	else
		teste_va_v_apr6 = isnumeric(va_v_apr6)
	end if
	
		if va_v_apr7="" or isnull(va_v_apr7) then
		va_v_apr7 = 0
		p_v_apr7="vazio"
		teste_va_v_apr7= true
	else
		teste_va_v_apr7 = isnumeric(va_v_apr7)
	end if
		if va_v_apr8="" or isnull(va_v_apr8) then
		va_v_apr8 = 0
		p_v_apr8="vazio"
		teste_va_v_apr8= true
	else
		teste_va_v_apr8 = isnumeric(va_v_apr8)
	end if



if teste_va_v_apr1=true and teste_va_v_apr2=true and teste_va_v_apr3=true and teste_va_v_apr4=true and teste_va_v_apr5=true and teste_va_v_apr6=true and teste_va_v_apr7=true and teste_va_v_apr8=true then
va_v_apr1=va_v_apr1*1
va_v_apr2=va_v_apr2*1
va_v_apr3=va_v_apr3*1
va_v_apr4=va_v_apr4*1
va_v_apr5=va_v_apr5*1
va_v_apr6=va_v_apr6*1
va_v_apr7=va_v_apr7*1
va_v_apr8=va_v_apr8*1



sum_p = va_v_apr1+va_v_apr2+va_v_apr3+va_v_apr4+va_v_apr5+va_v_apr6+va_v_apr7+va_v_apr8
'response.Write(">>"&sum_p&"<BR>")

	if sum_p>100 then
			fail = 1 
			erro = "sp"
			url = 0&"_"&sum_p&"_"&erro
			grava = "no"
	end if

else
			fail = 1 
			erro = "pt"
			url = 0&"_"&sum_p&"_"&erro
			grava = "no"
end if

'///////////////////////////////////////////////////////////////////////////

'APRs
	if va_apr1="" or isnull(va_apr1) then
		va_apr1=0
		s_va_apr1=0
	else
		teste_va_apr1 = isnumeric(va_apr1)
		if teste_va_apr1= true then
			'if p_v_apr1="vazio" then
			'fail = 1 
			'erro = "pv1"
			'url = 0&"_ _"&erro
			'grava = "no"					
			if teste_va_v_apr1 = true then
va_apr1=va_apr1*1
va_v_apr1=va_v_apr1*1			
				if va_apr1 =< va_v_apr1 then
					if va_apr1 =<100 then
					s_va_apr1=va_apr1
					
					else
					fail = 1 
					erro = "n1"
					url = nu_matricula&"_"&va_apr1&"_"&erro
					grava = "no"
					end if				
				else
				fail = 1 
				erro = "np1"
				url = nu_matricula&"_"&va_apr1&"_"&erro
				grava = "no"
				end if
			else
			fail = 1 
			erro = "n1"
			url = nu_matricula&"_"&va_apr1&"_"&erro
			grava = "no"

			end if
		else
			fail = 1 
			erro = "p1"
			url = nu_matricula&"_"&va_apr1&"_"&erro
			grava = "no"
		end if
	end if

	if va_apr2="" or isnull(va_apr2) then
		va_apr2=0
		s_va_apr2=0
	else
		teste_va_apr2 = isnumeric(va_apr2)
		if teste_va_apr2= true then
			'if p_v_apr2="vazio" then
			'fail = 1 
			'erro = "pv2"
			'url = 0&"_ _"&erro
			'grava = "no"					
			if teste_va_v_apr2 = true then

va_apr2=va_apr2*1
va_v_apr2=va_v_apr2*1			
				if va_apr2 =< va_v_apr2 then
					if va_apr2=<100 then
					s_va_apr2=va_apr2
					
					else
					fail = 1 
					erro = "n2"
					url = nu_matricula&"_"&va_apr2&"_"&erro
					grava = "no"
					end if				
				else
				fail = 1 
				erro = "np2"
				url = nu_matricula&"_"&va_apr2&"_"&erro
				grava = "no"
				end if

			else
				fail = 1 
				erro = "n2"
				url = nu_matricula&"_"&va_apr2&"_"&erro
				grava = "no"
			end if
		else
			fail = 1 
			erro = "p2"
			url = nu_matricula&"_"&va_apr2&"_"&erro
			grava = "no"
		end if
	end if

	if va_apr3="" or isnull(va_apr3) then
		va_apr3=0
		s_va_apr3=0
	else
		teste_va_apr3 = isnumeric(va_apr3)
		if teste_va_apr3= true then
			'if p_v_apr3="vazio" then
			'fail = 1 
			'erro = "pv3"
			'url = 0&"_ _"&erro
			'grava = "no"					
			if teste_va_v_apr3 = true then

va_apr3=va_apr3*1
va_v_apr3=va_v_apr3*1			
				if va_apr3 =< va_v_apr3 then
					if va_apr3=<100 then
					s_va_apr3=va_apr3
					
					else
					fail = 1 
					erro = "n3"
					url = nu_matricula&"_"&va_apr3&"_"&erro
					grava = "no"
					end if				
				else
				fail = 1 
				erro = "np3"
				url = nu_matricula&"_"&va_apr3&"_"&erro
				grava = "no"
				end if

			else
				fail = 1 
				erro = "n3"
				url = nu_matricula&"_"&va_apr3&"_"&erro
				grava = "no"
			end if
		else
			fail = 1 
			erro = "p3"
			url = nu_matricula&"_"&va_apr3&"_"&erro
			grava = "no"
		end if
	end if

	if va_apr4="" or isnull(va_apr4) then
		va_apr4=0
		s_va_apr4=0
	else
		teste_va_apr4 = isnumeric(va_apr4)
		if teste_va_apr4= true then
			'if p_v_apr4="vazio" then
			'fail = 1 
			'erro = "pv4"
			'url = 0&"_ _"&erro
			'grava = "no"					
			if teste_va_v_apr4 = true then
va_apr4=va_apr4*1
va_v_apr4=va_v_apr4*1			
				if va_apr4 =< va_v_apr4 then
					if va_apr4=<100 then
					s_va_apr4=va_apr4
					
					else
					fail = 1 
					erro = "n4"
					url = nu_matricula&"_"&va_apr4&"_"&erro
					grava = "no"
					end if				
				else
				fail = 1 
				erro = "np4"
				url = nu_matricula&"_"&va_apr4&"_"&erro
				grava = "no"
				end if
			else
				fail = 1 
				erro = "n4"
				url = nu_matricula&"_"&va_apr4&"_"&erro
				grava = "no"
			end if
		else
			fail = 1 
			erro = "p4"
			url = nu_matricula&"_"&va_apr4&"_"&erro
			grava = "no"
		end if
	end if

	if va_apr5="" or isnull(va_apr5) then
		va_apr5=0
		s_va_apr5=0
	else
		teste_va_apr5 = isnumeric(va_apr5)


		if teste_va_apr5= true then
			'if p_v_apr5="vazio" then
			'fail = 1 
			'erro = "pv5"
			'url = 0&"_ _"&erro
			'grava = "no"					
			if teste_va_v_apr5 = true then
va_apr5=va_apr5*1
va_v_apr5=va_v_apr5*1			
				if va_apr5 =< va_v_apr5 then
					if va_apr5=<100 then
					s_va_apr5=va_apr5
					
					else
					fail = 1 
					erro = "n5"
					url = nu_matricula&"_"&va_apr5&"_"&erro
					grava = "no"
					end if				
				else
				fail = 1 
				erro = "np5"
				url = nu_matricula&"_"&va_apr5&"_"&erro
				grava = "no"
				end if
			else
				fail = 1 
				erro = "n5"
				url = nu_matricula&"_"&va_apr5&"_"&erro
				grava = "no"
			end if
		else
			fail = 1 
			erro = "p5"
			url = nu_matricula&"_"&va_apr5&"_"&erro
			grava = "no"
		end if
	end if

	if va_apr6="" or isnull(va_apr6) then
		va_apr6=0
		s_va_apr6=0
	else
		teste_va_apr6 = isnumeric(va_apr6)
		if teste_va_apr6= true then
			'if p_v_apr6="vazio" then
			'fail = 1 
			'erro = "pv6"
			'url = 0&"_ _"&erro
			'grava = "no"					
			if teste_va_v_apr6 = true then
va_apr6=va_apr6*1
va_v_apr6=va_v_apr6*1			
				if va_apr6 =< va_v_apr6 then
					if va_apr6=<100 then
					s_va_apr6=va_apr6
					
					else
					fail = 1 
					erro = "n6"
					url = nu_matricula&"_"&va_apr6&"_"&erro
					grava = "no"
					end if				
				else
				fail = 1 
				erro = "np6"
				url = nu_matricula&"_"&va_apr6&"_"&erro
				grava = "no"
				end if
			else
				fail = 1 
				erro = "n6"
				url = nu_matricula&"_"&va_apr6&"_"&erro
				grava = "no"
			end if
		else
			fail = 1 
			erro = "p6"
			url = nu_matricula&"_"&va_apr6&"_"&erro
			grava = "no"
		end if
	end if
	
	if va_apr7="" or isnull(va_apr7) then
		va_apr7=0
		s_va_apr7=0
	else
		teste_va_apr7 = isnumeric(va_apr7)
		if teste_va_apr7= true then
			'if p_v_apr7="vazio" then
			'fail = 1 
			'erro = "pv7"
			'url = 0&"_ _"&erro
			'grava = "no"					
			if teste_va_v_apr7 = true then
va_apr7=va_apr7*1
va_v_apr7=va_v_apr7*1			
				if va_apr7 =< va_v_apr7 then
					if va_apr7=<100 then
					s_va_apr7=va_apr7
					
					else
					fail = 1 
					erro = "n7"
					url = nu_matricula&"_"&va_apr7&"_"&erro
					grava = "no"
					end if				
				else
				fail = 1 
				erro = "np7"
				url = nu_matricula&"_"&va_apr7&"_"&erro
				grava = "no"
				end if
			else
				fail = 1 
				erro = "n7"
				url = nu_matricula&"_"&va_apr7&"_"&erro
				grava = "no"
			end if
		else
			fail = 1 
			erro = "p7"
			url = nu_matricula&"_"&va_apr7&"_"&erro
			grava = "no"
		end if
	end if
	
	if va_apr8="" or isnull(va_apr8) then
		va_apr8=0
		s_va_apr8=0
	else
		teste_va_apr8 = isnumeric(va_apr8)
		if teste_va_apr8= true then
			'if p_v_apr6="vazio" then
			'fail = 1 
			'erro = "pv6"
			'url = 0&"_ _"&erro
			'grava = "no"					
			if teste_va_v_apr8 = true then
va_apr8=va_apr8*1
va_v_apr8=va_v_apr8*1			
				if va_apr8 =< va_v_apr8 then
					if va_apr8=<100 then
					s_va_apr8=va_apr8
					
					else
					fail = 1 
					erro = "n8"
					url = nu_matricula&"_"&va_apr8&"_"&erro
					grava = "no"
					end if				
				else
				fail = 1 
				erro = "np8"
				url = nu_matricula&"_"&va_apr8&"_"&erro
				grava = "no"
				end if
			else
				fail = 1 
				erro = "n8"
				url = nu_matricula&"_"&va_apr8&"_"&erro
				grava = "no"
			end if
		else
			fail = 1 
			erro = "p8"
			url = nu_matricula&"_"&va_apr8&"_"&erro
			grava = "no"
		end if
	end if	
	

'response.Write(i&"-"&nu_matricula&"-"&va_apr8 &"=<"& va_v_apr8&"<BR>")

'//////////////////////////////////////////////////////////////////////
'Notas

	if va_pr="" or isnull(va_pr) then
		va_pr=0
		s_va_pr=0
	else
		teste_va_pr = isnumeric(va_pr)
		if teste_va_pr = true then
			if va_pr>=0 and va_pr<=100 then
				s_va_pr=va_pr	
			else
				fail = 1 
				erro = "pr"
				url = nu_matricula&"_"&va_pr&"_"&erro
				grava = "no"				
			end if
			
		else
			fail = 1 
			erro = "pr"
			url = nu_matricula&"_"&va_pr&"_"&erro
			grava = "no"
		end if
	
	end if

	if va_te="" or isnull(va_te) then
		va_te=0
		s_va_te=0
	else
		teste_va_te = isnumeric(va_te)
		if teste_va_te = true then
			if va_te>=0 and va_te<=100 then
				s_va_te=va_te
			else
				fail = 1 
				erro = "te"
				url = nu_matricula&"_"&va_te&"_"&erro
				grava = "no"				
			end if

		else
			fail = 1 
			erro = "te"
			url = nu_matricula&"_"&va_te&"_"&erro
			grava = "no"
		end if
	end if
va_pr=va_pr*1
va_te=va_te*1
pr1pr2=va_pr+va_te

	if pr1pr2>10 then
			fail = 1 
			erro = "pr1pr2"
			url = nu_matricula&"_"&pr1pr2&"_"&erro
			grava = "no"
	else
	pr1pr2=pr1pr2
	end if

	if va_bon="" or isnull(va_bon) then
		va_bon=0
		s_va_bon=0
	else
		teste_va_bon = isnumeric(va_bon) 
		if teste_va_bon = true then
			if va_bon>=0 and va_bon<=10 then
				s_va_bon=va_bon
			else
				fail = 1 
				erro = "bon"
				url = nu_matricula&"_"&va_bon&"_"&erro
				grava = "no"				
			end if


		else
			fail = 1 
			erro = "bon"
			url = nu_matricula&"_"&va_bon&"_"&erro
			grava = "no"
		end if
	end if

'/////////////////////////////////////////////////////////////////////////
'Médias


if grava = "ok" then
va_apr1=s_va_apr1*1
va_apr2=s_va_apr2*1
va_apr3=s_va_apr3*1
va_apr4=s_va_apr4*1
va_apr5=s_va_apr5*1
va_apr6=s_va_apr6*1
va_apr7=s_va_apr7*1
va_apr8=s_va_apr8*1


va_sapr = va_apr1+va_apr2+va_apr3+va_apr4+va_apr5+va_apr6+va_apr7+va_apr8


	decimo = va_sapr - Int(va_sapr)
		If decimo >= 0.5 Then
			nota_arredondada = Int(va_sapr) + 1
			va_sapr=nota_arredondada
		Else
			nota_arredondada = Int(va_sapr)
			va_sapr=nota_arredondada					
		End If
		va_sapr=va_sapr/10	
	va_sapr = formatNumber(va_sapr,1)


va_me = (va_sapr+pr1pr2)/2

	va_me=va_me*10
	decimo = va_me - Int(va_me)
		If decimo >= 0.5 Then
			nota_arredondada = Int(va_me) + 1
			va_me=nota_arredondada
		Else
			nota_arredondada = Int(va_me)
			va_me=nota_arredondada					
		End If

	va_me = formatNumber(va_me,1)


if grava = "ok" then

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		CONEXAO0 = "Select * from TB_Nota_C WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"'"
		Set RS0 = CON.Execute(CONEXAO0)
		
If RS0.EOF THEN	


	if periodo=1 then
		va_me=va_me*1
		s_va_bon=s_va_bon*1
		va_mc = va_me + s_va_bon
		va_me=va_mc		
		if va_mc= "" or isnull(va_mc) then
		va_mc=0 
		end if
		va_mc2 = va_mc/2
		va_mc3 = va_mc2/2
		va_mc4 = va_mc3/2
	elseif periodo=2 then
		va_mc=va_mc*1
		va_me=va_me*1
		s_va_bon=s_va_bon*10

		if va_mc= "" or isnull(va_mc) then
			va_mc=0 
		end if
			
		if ((va_mc + va_me)/2) >= 70 then
			va_mc2=((media1 + va_me)/2)		 
		elseif ((((va_mc + va_me)/2) + s_va_bon)/2)> ((va_mc + va_me)/2) then
			
			va_mc2=((((va_mc + va_me)/2) + s_va_bon)/2)
			if va_mc2>70 then
				va_mc2=70
			end if
		else
			va_mc2=((va_mc + va_me)/2) 

		end if					
		va_mc3 = va_mc2/2
		va_mc4 = va_mc3/2
	elseif periodo=3 then
		va_me=va_me*1
		s_va_bon=s_va_bon*1
		va_mc = (va_me/2)+ s_va_bon
	if va_mc= "" or isnull(va_mc) then
	va_mc=0 
	end if	
	va_mc4 = va_mc/2
	elseif periodo=4 then
		va_me=va_me*1
		va_mc = va_me/2
	if va_mc= "" or isnull(va_mc) then
	va_mc=0 
	end if		
	end if

va_mc=va_mc*10
		decimo = va_mc - Int(va_mc)
		If decimo >= 0.5 Then
			nota_arredondada = Int(va_mc) + 1
			va_mc=nota_arredondada
		Else
			nota_arredondada = Int(va_mc)
			va_mc=nota_arredondada					
		End If
	va_mc=va_mc/10	
	va_mc = formatNumber(va_mc,1)
	


	va_mc2=va_mc2*10
		decimo = va_mc2 - Int(va_mc2)
		If decimo >= 0.5 Then
			nota_arredondada = Int(va_mc2) + 1
			va_mc2=nota_arredondada
		Else
			nota_arredondada = Int(va_mc2)
			va_mc2=nota_arredondada					
		End If
	va_mc2=va_mc2/10
va_mc2 = formatNumber(va_mc2,1)

	va_mc3=va_mc3*10
		decimo = va_mc3 - Int(va_mc3)
		If decimo >= 0.5 Then
			nota_arredondada = Int(va_mc3) + 1
			va_mc3=nota_arredondada
		Else
			nota_arredondada = Int(va_mc3)
			va_mc3=nota_arredondada					
		End If
	va_mc3=va_mc3/10
va_mc3 = formatNumber(va_mc3,1)

	va_mc4=va_mc4*10
		decimo = va_mc4 - Int(va_mc4)
		If decimo >= 0.5 Then
			nota_arredondada = Int(va_mc4) + 1
			va_mc4=nota_arredondada
		Else
			nota_arredondada = Int(va_mc4)
			va_mc4=nota_arredondada					
		End If
	va_mc4=va_mc4/10
va_mc4 = formatNumber(va_mc4,1)


'response.Write("4"&turma &"/"&co_materia_pr)
Set RS = server.createobject("adodb.recordset")

RS.open "TB_Nota_C", CON, 2, 2 'which table do you want open
RS.addnew

if periodo=1 then
	RS("CO_Matricula") = nu_matricula
	RS("CO_Materia_Principal") = co_materia_pr
	RS("CO_Materia") = co_materia
	RS("Apr1_P1")=va_apr1
	RS("Apr2_P1")=va_apr2
	RS("Apr3_P1")=va_apr3
	RS("Apr4_P1")=va_apr4
	RS("Apr5_P1")=va_apr5
	RS("Apr6_P1")=va_apr6
	RS("Apr7_P1")=va_apr7
	RS("Apr8_P1")=va_apr8	
	RS("V_Apr1_P1")=va_v_apr1
	RS("V_Apr2_P1")=va_v_apr2
	RS("V_Apr3_P1")=va_v_apr3
	RS("V_Apr4_P1")=va_v_apr4
	RS("V_Apr5_P1")=va_v_apr5
	RS("V_Apr6_P1")=va_v_apr6
	RS("V_Apr7_P1")=va_v_apr7
	RS("V_Apr8_P1")=va_v_apr8
	RS("VA_Sapr1")=va_sapr
	RS("VA_Pr1")=va_pr
	RS("VA_Te1")=va_te
	RS("VA_Bon1")=va_bon
	RS("VA_Me1")=va_me
	RS("VA_Mc1")=va_mc
	RS("NU_Faltas_P1")=va_faltas
	RS("VA_Mc2")=va_mc2
	RS("VA_Mc3")=va_mc3
	RS("VA_Mc4")=va_mc4
	RS("DA_Ult_Acesso") = data
	RS("HO_ult_Acesso") = horario
	RS("CO_Usuario")= co_usr
elseif periodo=2 then
	RS("CO_Matricula") = nu_matricula
	RS("CO_Materia_Principal") = co_materia_pr
	RS("CO_Materia") = co_materia
	RS("Apr1_P2")=va_apr1
	RS("Apr2_P2")=va_apr2
	RS("Apr3_P2")=va_apr3
	RS("Apr4_P2")=va_apr4
	RS("Apr5_P2")=va_apr5
	RS("Apr6_P2")=va_apr6
	RS("Apr7_P2")=va_apr7
	RS("Apr8_P2")=va_apr8	
	RS("V_Apr1_P2")=va_v_apr1
	RS("V_Apr2_P2")=va_v_apr2
	RS("V_Apr3_P2")=va_v_apr3
	RS("V_Apr4_P2")=va_v_apr4
	RS("V_Apr5_P2")=va_v_apr5
	RS("V_Apr6_P2")=va_v_apr6
	RS("V_Apr7_P2")=va_v_apr7
	RS("V_Apr8_P2")=va_v_apr8	
	RS("VA_Sapr2")=va_sapr
	RS("VA_Pr2")=va_pr
	RS("VA_Te2")=va_te
	RS("VA_Bon2")=va_bon
	RS("VA_Me2")=va_me
	RS("VA_Mc2")=va_mc
	RS("NU_Faltas_P2")=va_faltas
	RS("VA_Mc3")=va_mc3
	RS("VA_Mc4")=va_mc4
	RS("DA_Ult_Acesso") = data
	RS("HO_ult_Acesso") = horario
	RS("CO_Usuario")= co_usr
elseif periodo=3 then
	RS("CO_Matricula") = nu_matricula
	RS("CO_Materia_Principal") = co_materia_pr
	RS("CO_Materia") = co_materia
	RS("Apr1_P3")=va_apr1
	RS("Apr2_P3")=va_apr2
	RS("Apr3_P3")=va_apr3
	RS("Apr4_P3")=va_apr4
	RS("Apr5_P3")=va_apr5
	RS("Apr6_P3")=va_apr6
	RS("Apr7_P3")=va_apr7
	RS("Apr8_P3")=va_apr8	
	RS("V_Apr1_P3")=va_v_apr1
	RS("V_Apr2_P3")=va_v_apr2
	RS("V_Apr3_P3")=va_v_apr3
	RS("V_Apr4_P3")=va_v_apr4
	RS("V_Apr5_P3")=va_v_apr5
	RS("V_Apr6_P3")=va_v_apr6
	RS("V_Apr7_P3")=va_v_apr7
	RS("V_Apr8_P3")=va_v_apr8	
	RS("VA_Sapr3")=va_sapr
	RS("VA_Pr3")=va_pr
	RS("VA_Te3")=va_te
	RS("VA_Bon3")=va_bon
	RS("VA_Me3")=va_me
	RS("VA_Mc3")=va_mc
	RS("NU_Faltas_P3")=va_faltas
	RS("VA_Mc4")=va_mc4
	RS("DA_Ult_Acesso") = data
	RS("HO_ult_Acesso") = horario
	RS("CO_Usuario")= co_usr
elseif periodo=4 then
	RS("CO_Matricula") = nu_matricula
	RS("CO_Materia_Principal") = co_materia_pr
	RS("CO_Materia") = co_materia
	RS("Apr1_EC")=va_apr1
	RS("Apr2_EC")=va_apr2
	RS("Apr3_EC")=va_apr3
	RS("Apr4_EC")=va_apr4
	RS("Apr5_EC")=va_apr5
	RS("Apr6_EC")=va_apr6
	RS("Apr7_EC")=va_apr7
	RS("Apr8_EC")=va_apr8
	RS("V_Apr7_EC")=va_v_apr7
	RS("V_Apr8_EC")=va_v_apr8		
	RS("V_Apr1_EC")=va_v_apr1
	RS("V_Apr2_EC")=va_v_apr2
	RS("V_Apr3_EC")=va_v_apr3
	RS("V_Apr4_EC")=va_v_apr4
	RS("V_Apr5_EC")=va_v_apr5
	RS("V_Apr6_EC")=va_v_apr6
	RS("VA_Sapr_EC")=va_sapr
	RS("VA_Pr4")=va_pr
	RS("VA_Me_EC")=va_me
	RS("VA_Mfinal")=va_mc
	RS("DA_Ult_Acesso") = data
	RS("HO_ult_Acesso") = horario
	RS("CO_Usuario")= co_usr
end if

RS.update
set RS=nothing

else
		media1=RS0("VA_Mc1")
		me2=RS0("VA_Me2")
		bon2=RS0("VA_Bon2")
		media2=RS0("VA_Mc2")
		me3=RS0("VA_Me3")
		bon3=RS0("VA_Bon3")
		media3=RS0("VA_Mc3")
		me4=RS0("VA_Me_EC")

if media1= "" or isnull(media1) then
media1=0
else
media1=media1*1
end if

if me2= "" or isnull(me2) then
me2=0
else
me2=me2*1
end if

if bon2= "" or isnull(bon2) then
bon2=0
else
bon2=bon2*1
end if

if media2= "" or isnull(media2) then
media2=0
else
media2=media2*1
end if

if me3= "" or isnull(me3) then
me3=0
else
me3=me3*1
end if

if bon3= "" or isnull(bon3) then
bon3=0
else
bon3=bon3*1
end if

if media3= "" or isnull(media3) then
media3=0
else
media3=media3*1
end if

if me4= "" or isnull(me4) then
me4=0
else
me4=me4*1
end if

	if periodo=1 then
		va_me=va_me*1
		s_va_bon=s_va_bon*1
		va_mc = va_me+s_va_bon
		va_me=va_mc		
'response.Write(va_mc &"="& va_me&"+"&s_va_bon&"-"&nu_matricula&"<BR>")
		va_mc2 = ((va_mc+me2)/2)+ bon2
		va_mc3 = ((va_mc2+me3)/2)+ bon3
		va_mc4 = (va_mc3+me4)/2
	elseif periodo=2 then
		media1=media1*1
		va_me=va_me*1
		s_va_bon=s_va_bon*10
'response.Write(media1&"<<BR>")
'response.Write(va_me&"<<BR>")
'response.Write(s_va_bon&"<<BR>")
		
		if ((media1 + va_me)/2) >= 70 then
			va_mc=((media1 + va_me)/2)
'			response.Write("OK")			 
		elseif ((((media1 + va_me)/2) + s_va_bon)/2)> ((media1 + va_me)/2) then
			
			va_mc=((((media1 + va_me)/2) + s_va_bon)/2)
			if va_mc>70 then
				va_mc=70
			end if
		else
			va_mc=((media1 + va_me)/2) 
		end if
		va_mc3 = ((va_mc+me3)/2)+ bon3
		va_mc4 = (va_mc3+me4)/2
'response.end()
	elseif periodo=3 then
		va_me=va_me*1
		s_va_bon=s_va_bon*1
		va_mc = ((media2+va_me)/2)+ s_va_bon
		va_mc4 = (va_mc+me4)/2
	elseif periodo=4 then
		va_me=va_me*1
		va_mc = (media3+va_me)/2
	end if

		va_mc=va_mc*10
		decimo = va_mc - Int(va_mc)
		If decimo >= 0.5 Then
			nota_arredondada = Int(va_mc) + 1
			va_mc=nota_arredondada
		Else
			nota_arredondada = Int(va_mc)
			va_mc=nota_arredondada					
		End If
	va_mc=va_mc/10	
	va_mc = formatNumber(va_mc,1)


	va_mc2=va_mc2*10
		decimo = va_mc2 - Int(va_mc2)
		If decimo >= 0.5 Then
			nota_arredondada = Int(va_mc2) + 1
			va_mc2=nota_arredondada
		Else
			nota_arredondada = Int(va_mc2)
			va_mc2=nota_arredondada					
		End If
	va_mc2=va_mc2/10
va_mc2 = formatNumber(va_mc2,1)

	va_mc3=va_mc3*10
		decimo = va_mc3 - Int(va_mc3)
		If decimo >= 0.5 Then
			nota_arredondada = Int(va_mc3) + 1
			va_mc3=nota_arredondada
		Else
			nota_arredondada = Int(va_mc3)
			va_mc3=nota_arredondada					
		End If
	va_mc3=va_mc3/10
va_mc3 = formatNumber(va_mc3,1)

	va_mc4=va_mc4*10
		decimo = va_mc4 - Int(va_mc4)
		If decimo >= 0.5 Then
			nota_arredondada = Int(va_mc4) + 1
			va_mc4=nota_arredondada
		Else
			nota_arredondada = Int(va_mc4)
			va_mc4=nota_arredondada					
		End If
	va_mc4=va_mc4/10
va_mc4 = formatNumber(va_mc4,1)
					

'if periodo=1 then
'teste= "UPDATE TB_Nota_C SET "&no_apr1&" ='"&va_apr1&"', "&no_apr2&" ='"&va_apr2&"', "&no_apr3&" ='"&va_apr3&"', "&no_apr4&" ='"&va_apr4&"', "&no_apr5&" ='"&va_apr5&"', "
'teste=teste&no_apr6&" ='"&va_apr6&"', "&no_v_apr1&" ='"&va_v_apr1&"', "&no_v_apr2&" ='"&va_v_apr2&"', "&no_v_apr3&" ='"&va_v_apr3&"', "&no_v_apr4&" ='"&va_v_apr4&"', "
'teste=teste&no_v_apr5&" ='"&va_v_apr5&"', "&no_v_apr6&" ='"&va_v_apr6&"', "&no_sapr&" ='"&va_sapr&"', "&no_pr&" ='"&va_pr&"', "&no_te&" ='"&va_te&"', "&no_bon&" ='"&va_bon&"',"&no_mc&" ='"&va_mc&"', "
'teste=teste&no_me&" ='"&va_me&"', "&no_faltas&" = '"&faltas&"', VA_Mc2='"&va_mc2&"', VA_Mc3='"&va_mc3&"', VA_Mfinal='"&va_mc4&"', DA_Ult_Acesso ='"& data &"', HO_ult_Acesso ='"& horario &"', CO_Usuario="& co_usr &"  WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"'"


'response.Write(teste&"<BR>")

'end if
'response.Write("5>>>>>"&no_v_apr7&" ='"&va_v_apr7&"-"&no_v_apr8&" ='"&va_v_apr8)
if periodo=1 then
	sql_atualiza= "UPDATE TB_Nota_C SET "&variaveis_bd(0)&" ='"&va_apr1&"', "&variaveis_bd(1)&" ='"&va_apr2&"', "&variaveis_bd(2)&" ='"&va_apr3&"', "&variaveis_bd(3)&" ='"
	sql_atualiza=sql_atualiza&va_apr4&"', "&variaveis_bd(4)&" ='"&va_apr5&"', "&variaveis_bd(5)&" ='"&va_apr6&"',"&variaveis_bd(6)&" ='"&va_apr7&"', "&variaveis_bd(7)&" ='"
	sql_atualiza=sql_atualiza&va_apr8&"', "&linha_pesos_variaveis(2)&" ='"&va_v_apr1&"', "&linha_pesos_variaveis(3)&" ='"&va_v_apr2&"', "&linha_pesos_variaveis(4)&" ='"
	sql_atualiza=sql_atualiza&va_v_apr3&"', "&linha_pesos_variaveis(5)&" ='"&va_v_apr4&"', "&linha_pesos_variaveis(6)&" ='"&va_v_apr5&"', "
	sql_atualiza=sql_atualiza&linha_pesos_variaveis(7)&" ='"&va_v_apr6&"', "&linha_pesos_variaveis(8)&" ='"&va_v_apr7&"', "&linha_pesos_variaveis(9)&" ='"&va_v_apr8&"', "
	sql_atualiza=sql_atualiza&variaveis_bd(8)&" ='"&va_sapr&"', "&variaveis_bd(9)&" ='"&va_pr&"', "&variaveis_bd(10)&" ='"&va_te&"', "&variaveis_bd(11)&" ='"&va_bon&"',"
	sql_atualiza=sql_atualiza&variaveis_bd(12)&" ='"&va_me&"', "&variaveis_bd(13)&" ='"&va_mc&"', "&variaveis_bd(14)&" = '"&faltas&"', VA_Mc2='"&va_mc2&"', VA_Mc3='"&va_mc3
	sql_atualiza=sql_atualiza&"', VA_Mfinal='"&va_mc4&"', DA_Ult_Acesso ='"& data &"', HO_ult_Acesso ='"& horario &"', CO_Usuario="& co_usr &"  WHERE CO_Matricula = "
	sql_atualiza=sql_atualiza& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"'"
elseif periodo=2 then
	sql_atualiza= "UPDATE TB_Nota_C SET "&variaveis_bd(0)&" ='"&va_apr1&"', "&variaveis_bd(1)&" ='"&va_apr2&"', "&variaveis_bd(2)&" ='"&va_apr3&"', "&variaveis_bd(3)&" ='"
	sql_atualiza=sql_atualiza&va_apr4&"', "&variaveis_bd(4)&" ='"&va_apr5&"', "&variaveis_bd(5)&" ='"&va_apr6&"',"&variaveis_bd(6)&" ='"&va_apr7&"', "&variaveis_bd(7)&" ='"
	sql_atualiza=sql_atualiza&va_apr8&"', "&linha_pesos_variaveis(2)&" ='"&va_v_apr1&"', "&linha_pesos_variaveis(3)&" ='"&va_v_apr2&"', "&linha_pesos_variaveis(4)&" ='"
	sql_atualiza=sql_atualiza&va_v_apr3&"', "&linha_pesos_variaveis(5)&" ='"&va_v_apr4&"', "&linha_pesos_variaveis(6)&" ='"&va_v_apr5&"', "
	sql_atualiza=sql_atualiza&linha_pesos_variaveis(7)&" ='"&va_v_apr6&"', "&linha_pesos_variaveis(8)&" ='"&va_v_apr7&"', "&linha_pesos_variaveis(9)&" ='"&va_v_apr8&"', "
	sql_atualiza=sql_atualiza&variaveis_bd(8)&" ='"&va_sapr&"', "&variaveis_bd(9)&" ='"&va_pr&"', "&variaveis_bd(10)&" ='"&va_te&"', "&variaveis_bd(11)&" ='"&va_bon&"',"
	sql_atualiza=sql_atualiza&variaveis_bd(12)&" ='"&va_me&"', "&variaveis_bd(13)&" ='"&va_mc&"', "&variaveis_bd(14)&" = '"&faltas&"', VA_Mc3='"&va_mc3
	sql_atualiza=sql_atualiza&"', VA_Mfinal='"&va_mc4&"', DA_Ult_Acesso ='"& data &"', HO_ult_Acesso ='"& horario &"', CO_Usuario="& co_usr &"  WHERE CO_Matricula = "
	sql_atualiza=sql_atualiza& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"'"
elseif periodo=3 then
	sql_atualiza= "UPDATE TB_Nota_C SET "&variaveis_bd(0)&" ='"&va_apr1&"', "&variaveis_bd(1)&" ='"&va_apr2&"', "&variaveis_bd(2)&" ='"&va_apr3&"', "&variaveis_bd(3)&" ='"
	sql_atualiza=sql_atualiza&va_apr4&"', "&variaveis_bd(4)&" ='"&va_apr5&"', "&variaveis_bd(5)&" ='"&va_apr6&"',"&variaveis_bd(6)&" ='"&va_apr7&"', "&variaveis_bd(7)&" ='"
	sql_atualiza=sql_atualiza&va_apr8&"', "&linha_pesos_variaveis(2)&" ='"&va_v_apr1&"', "&linha_pesos_variaveis(3)&" ='"&va_v_apr2&"', "&linha_pesos_variaveis(4)&" ='"
	sql_atualiza=sql_atualiza&va_v_apr3&"', "&linha_pesos_variaveis(5)&" ='"&va_v_apr4&"', "&linha_pesos_variaveis(6)&" ='"&va_v_apr5&"', "
	sql_atualiza=sql_atualiza&linha_pesos_variaveis(7)&" ='"&va_v_apr6&"', "&linha_pesos_variaveis(8)&" ='"&va_v_apr7&"', "&linha_pesos_variaveis(9)&" ='"&va_v_apr8&"', "
	sql_atualiza=sql_atualiza&variaveis_bd(8)&" ='"&va_sapr&"', "&variaveis_bd(9)&" ='"&va_pr&"', "&variaveis_bd(10)&" ='"&va_te&"', "&variaveis_bd(11)&" ='"&va_bon&"',"
	sql_atualiza=sql_atualiza&variaveis_bd(12)&" ='"&va_me&"', "&variaveis_bd(13)&" ='"&va_mc&"', "&variaveis_bd(14)&" = '"&faltas
	sql_atualiza=sql_atualiza&"', VA_Mfinal='"&va_mc4&"', DA_Ult_Acesso ='"& data &"', HO_ult_Acesso ='"& horario &"', CO_Usuario="& co_usr &"  WHERE CO_Matricula = "
	sql_atualiza=sql_atualiza& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"'"	
	
else
	sql_atualiza= "UPDATE TB_Nota_C SET "&variaveis_bd(0)&" ='"&va_apr1&"', "&variaveis_bd(1)&" ='"&va_apr2&"', "&variaveis_bd(2)&" ='"&va_apr3&"', "&variaveis_bd(3)&" ='"
	sql_atualiza=sql_atualiza&va_apr4&"', "&variaveis_bd(4)&" ='"&va_apr5&"', "&variaveis_bd(5)&" ='"&va_apr6&"',"&variaveis_bd(6)&" ='"&va_apr7&"', "&variaveis_bd(7)&" ='"
	sql_atualiza=sql_atualiza&va_apr8&"', "&linha_pesos_variaveis(2)&" ='"&va_v_apr1&"', "&linha_pesos_variaveis(3)&" ='"&va_v_apr2&"', "&linha_pesos_variaveis(4)&" ='"
	sql_atualiza=sql_atualiza&va_v_apr3&"', "&linha_pesos_variaveis(5)&" ='"&va_v_apr4&"', "&linha_pesos_variaveis(6)&" ='"&va_v_apr5&"', "
	sql_atualiza=sql_atualiza&linha_pesos_variaveis(7)&" ='"&va_v_apr6&"', "&linha_pesos_variaveis(8)&" ='"&va_v_apr7&"', "&linha_pesos_variaveis(9)&" ='"&va_v_apr8&"', "
	sql_atualiza=sql_atualiza&variaveis_bd(8)&" ='"&va_sapr&"', "&variaveis_bd(9)&" ='"&va_pr&"',"&variaveis_bd(10)&" ='"&va_me&"', "&variaveis_bd(11)&" ='"&va_mc&"', "
	sql_atualiza=sql_atualiza&"DA_Ult_Acesso ='"& data &"', HO_ult_Acesso ='"& horario &"', CO_Usuario="& co_usr &"  WHERE CO_Matricula = "& nu_matricula &" AND "	
	sql_atualiza=sql_atualiza&"CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"'"
end if
Set RS2 = Con.Execute(sql_atualiza)

end if
end if
end if
'response.Write(i&"-grava-"&grava&"hp=err_"&url&"&obr="&obr&"<br>"&sql_atualiza)
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