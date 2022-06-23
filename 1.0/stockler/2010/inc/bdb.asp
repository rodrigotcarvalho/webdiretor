<%On Error Resume Next%>
<!--#include file="caminhos.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="../../../global/funcoes_diversas.asp" -->
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
				av1=request.form("av1_"&i)
				av2=request.form("av2_"&i)
				av3=request.form("av3_"&i)
				pr=request.form("pr_"&i)
				at=request.form("at_"&i)
				va_rec=request.form("rec_"&i)
		end if	
	else
			nu_matricula = request.form("nu_matricula_"&i)			
			va_faltas=request.form("faltas_"&i)
			av1=request.form("av1_"&i)
			av2=request.form("av2_"&i)
			av3=request.form("av3_"&i)
			pr=request.form("pr_"&i)
			at=request.form("at_"&i)
			va_rec=request.form("rec_"&i)
	end if
	
if fail = 0 then 		
Session("faltas")=va_faltas
Session("av1")=av1
Session("av2")=av2
Session("av3")=av3
Session("pr")=pr
Session("at")=at
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
'AVALIAÇÕES
s_av1=0
if av1="" or isnull(av1) then
	av1=NULL		
	s_av1=0
	soma_av1=0	
	divisor_av1=0	
else
	teste_av1 = isnumeric(av1)
	if teste_av1= true then					
		av1=av1*1			
		if av1 =<10 then
			s_av1=1
			divisor_av1=1
			soma_av1=av1																									
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "av1$1"
				matric_Erro=i
				url = nu_matricula&"_"&av1&"_"&erro
				grava = "no"
			end if					
		end if				
	else
		if  fail = 1 then
			grava = "no"
		else
			fail = 1 
			erro = "av1$0"
			matric_Erro=i
			url = nu_matricula&"_"&av1&"_"&erro
			grava = "no"
		end if
	end if
end if
s_av2=0
if av2="" or isnull(av2) then
	av2=NULL		
	s_av2=0
	soma_av2=0		
	divisor_av2=0
else
	teste_av2 = isnumeric(av2)
	if teste_av2= true then					
		av2=av2*1			
		if av2 =<10 then
			s_av2=1
			divisor_av2=1
			soma_av2=av2																									
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "av2$1"
				matric_Erro=i
				url = nu_matricula&"_"&av2&"_"&erro
				grava = "no"
			end if					
		end if				
	else
		if  fail = 1 then
			grava = "no"
		else
			fail = 1 
			erro = "av2$0"
			matric_Erro=i
			url = nu_matricula&"_"&av2&"_"&erro
			grava = "no"
		end if
	end if
end if
s_av3=0
if av3="" or isnull(av3) then
	av3=NULL		
	s_av3=0
	soma_av3=0	
	divisor_av3=0	
else
	teste_av3 = isnumeric(av3)
	if teste_av3= true then					
		av3=av3*1			
		if av3 =<10 then
			s_av3=1
			divisor_av3=1
			soma_av3=av3																									
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "av3$1"
				matric_Erro=i
				url = nu_matricula&"_"&av3&"_"&erro
				grava = "no"
			end if					
		end if				
	else
		if  fail = 1 then
			grava = "no"
		else
				fail = 1 
				erro = "av3$0"
				matric_Erro=i
				url = nu_matricula&"_"&av3&"_"&erro
				grava = "no"
		end if
	end if
end if
'//////////////////////////////////////////////////////////////////////
s_pr=0
if pr="" or isnull(pr) then
	pr=NULL		
	s_pr=0
	soma_pr=0		
else
	teste_pr = isnumeric(pr)
	if teste_pr= true then					
		pr=pr*1			
		if pr =<10 then
			s_pr=1
			soma_pr=pr																									
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "pr$1"
				matric_Erro=i
				url = nu_matricula&"_"&pr&"_"&erro
				grava = "no"
			end if					
		end if				
	else
		if  fail = 1 then
			grava = "no"
		else
			fail = 1 
			erro = "pr$0"
			matric_Erro=i
			url = nu_matricula&"_"&pr&"_"&erro
			grava = "no"
		end if
	end if
end if


if at="" or isnull(at) then
	at=NULL		
	s_at=0
else
	teste_at = isnumeric(at) 
	if teste_at = true then
		at=at*1	
		if at=<10 then
			s_at=at																				
		else
			if  fail = 1 then
				grava = "no"
			else											
				fail = 1 
				erro = "at$1"
				url = nu_matricula&"_"&at&"_"&erro
				grava = "no"
			end if			
		end if
	else
		if  fail = 1 then
			grava = "no"
		else											
			fail = 1 
			erro = "at$0"
			url = nu_matricula&"_"&at&"_"&erro
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
		if va_rec=<10 then
			va_rec=va_rec*1
			s_va_rec=va_rec													
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
divisor_av1=divisor_av1*1
divisor_av2=divisor_av2*1
divisor_av3=divisor_av3*1
divisor_av=divisor_av1+divisor_av2+divisor_av3

s_va_t1=s_va_t1*1
s_va_t2=s_va_t2*1
s_va_t3=s_va_t3*1
soma_av1=soma_av1*1
soma_av2=soma_av2*1
soma_av3=soma_av3*1	

if s_av1=1 or s_av2=1 or s_av3=1 then
	media_teste=(soma_av1+soma_av2+soma_av3)/divisor_av
	media_t="ok"
	media_teste = arredonda(media_teste,"mat_dez",1,0)
else
	media_teste=NULL
	media_t="no"
end if

if media_t="ok" and s_pr=1 then
	m1=(media_teste+pr)/2					
	m1 = arredonda(m1,"mat_dez",1,0)
else
	m1=NULL
end if	
	
if isnull(m1) or m1="" then
	m2=NULL
	m3=NULL	
else		
	if isnull(at) or at="" then
		m2=m1		
	else
		m1=m1*1		
		at=at*1
		m2=m1+at
		m2 = arredonda(m2,"mat_dez",1,0)
		if m2>10 then
			if  fail = 1 then
				grava = "no"
			else											
				fail = 1 
				erro = "m2"
				url = nu_matricula&"_"&at&"_"&erro
				grava = "no"
			end if
		end if
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
			RS("VA_AV1")=av1
			RS("VA_AV2")=av2
			RS("VA_AV3")=av3
			RS("MAV_Avaliacao")=media_teste
			RS("VA_Prova")=pr						
			RS("VA_Media1")=m1
			RS("VA_Bonus")=at	
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
			RS("VA_AV1")=av1
			RS("VA_AV2")=av2
			RS("VA_AV3")=av3
			RS("MAV_Avaliacao")=media_teste
			RS("VA_Prova")=pr						
			RS("VA_Media1")=m1
			RS("VA_Bonus")=at	
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