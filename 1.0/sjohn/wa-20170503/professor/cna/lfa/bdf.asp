<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<%
call cabecalho(1)
chave = session("chave")
session("chave")=chave

opt=request.QueryString("opt")
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 
data = dia &"/"& mes &"/"& ano
horario = hora & ":"& min


unidade= request.form("unidade")
curso= request.form("curso")
etapa= request.form("etapa")
turma= request.form("turma")
ano_letivo = request.form("ano_letivo")
co_usr = request.form("co_usr")
max = request.form("max")

obr=unidade&"?"&curso&"?"&etapa&"?"&turma&"?"&ano_letivo

i=1

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_nf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CONMT = Server.CreateObject("ADODB.Connection") 
		ABRIRMT = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONMT.Open ABRIRMT
	
fail = 0
for i=1 to max
	grava="ok"

	
	nu_matricula = request.form("nu_matricula_"&i)
		
	if nu_matricula = "falta" then
		while nu_matricula = "falta" 
			i=i+1
			nu_matricula = request.form("nu_matricula_"&i)
		wend
	va_f1=request.form("f1_"&i)
	va_f2=request.form("f2_"&i)
	va_f3=request.form("f3_"&i)
	va_f4=request.form("f4_"&i)
	else
	nu_matricula = request.form("nu_matricula_"&i)
	va_f1=request.form("f1_"&i)
	va_f2=request.form("f2_"&i)
	va_f3=request.form("f3_"&i)
	va_f4=request.form("f4_"&i)
	end if
'response.Write(i&"-a<BR>")	
'response.Write(nu_matricula&"-a<BR>")	
if fail = 0 then 	
Session("va_f1")=va_f1
Session("va_f2")=va_f2
Session("va_f3")=va_f3
Session("va_f4")=va_f4	
end if	

'TESTES
s_va_t=0
	if va_f1="" or isnull(va_f1) then
		va_f1=NULL		
		s_va_f1=0
		soma_teste1=0		
	else
		teste_va_f1 = isnumeric(va_f1)
		if teste_va_f1= true then					
		va_f1=va_f1*1			
					if int(va_f1) =va_f1 then
					s_va_f1=1
					soma_teste1=va_f1															
					else
					fail = 1 
					erro = "f1"
					url = nu_matricula&"_"&va_f1&"_"&erro
					grava = "no"
					end if				
			else
			fail = 1 
			erro = "f1"
			url = nu_matricula&"_"&va_f1&"_"&erro
			grava = "no"

			end if
	end if

	if va_f2="" or isnull(va_f2) then
		va_f2=NULL		
		s_va_f2=0
		soma_teste2=0		
	else
		teste_va_f2 = isnumeric(va_f2)
		if teste_va_f2= true then					
		va_f2=va_f2*1			
					if int(va_f2) =va_f2 then
					s_va_f2=1
					soma_teste2=va_f2															
					else
					fail = 1 
					erro = "f2"
					url = nu_matricula&"_"&va_f2&"_"&erro
					grava = "no"
					end if				
			else
			fail = 1 
			erro = "f2"
			url = nu_matricula&"_"&va_f2&"_"&erro
			grava = "no"
			end if
	end if	
	
	
	if va_f3="" or isnull(va_f3) then
		va_f3=NULL		
		s_va_f3=0
		soma_teste3=0		
	else
		teste_va_f3 = isnumeric(va_f3)
		if teste_va_f3= true then					
		va_f3=va_f3*1			
					if int(va_f3) =va_f3 then
					s_va_f3=1
					soma_teste3=va_f3															
					else
					fail = 1 
					erro = "f3"
					url = nu_matricula&"_"&va_f3&"_"&erro
					grava = "no"
					end if				
			else
			fail = 1 
			erro = "f3"
			url = nu_matricula&"_"&va_f3&"_"&erro
			grava = "no"
			end if
	end if
	
	if va_f4="" or isnull(va_f4) then
		va_f4=NULL		
		s_va_f4=0
		soma_teste4=0		
	else
		teste_va_f4 = isnumeric(va_f4)
		if teste_va_f4= true then					
		va_f4=va_f4*1			
					if int(va_f4) =va_f4 then
					s_va_f4=1
					soma_teste4=va_f4															
					else
					fail = 1 
					erro = "f4"
					url = nu_matricula&"_"&va_f4&"_"&erro
					grava = "no"
					end if				
			else
			fail = 1 
			erro = "f4"
			url = nu_matricula&"_"&va_f4&"_"&erro
			grava = "no"
			end if
	end if	
	

if grava="no" then
else	
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		CONEXAO0 = "Select * from TB_Frequencia_Periodo WHERE CO_Matricula = "& nu_matricula 
		Set RS0 = CON.Execute(CONEXAO0)
		
	If RS0.EOF THEN	
	
		
		'response.Write("4"&turma &"/"&co_materia_pr)
		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Frequencia_Periodo", CON, 2, 2 'which table do you want open
		RS.addnew
			RS("CO_Matricula") = nu_matricula	
			RS("NU_Faltas_P1")=va_f1
			RS("NU_Faltas_P2")=va_f2
			RS("NU_Faltas_P3")=va_f3
			RS("NU_Faltas_P4")=va_f4
		
		RS.update
		set RS=nothing
		
	else
		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		CONEXAO0 = "DELETE * from TB_Frequencia_Periodo WHERE CO_Matricula = "& nu_matricula
		Set RS0 = CON.Execute(CONEXAO0)

		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Frequencia_Periodo", CON, 2, 2 'which table do you want open
		RS.addnew
		
			RS("CO_Matricula") = nu_matricula	
			RS("NU_Faltas_P1")=va_f1
			RS("NU_Faltas_P2")=va_f2
			RS("NU_Faltas_P3")=va_f3
			RS("NU_Faltas_P4")=va_f4
		
		RS.update
		set RS=nothing		
		
'		sql_atualiza= "UPDATE TB_Frequencia_Periodo SET VA_Teste1 ="&sql_va_f1&", VA_Teste2 ="&sql_va_f2&", VA_Teste3 ="&sql_va_f3&", VA_Teste4 ="&sql_va_f4&", MD_Teste =FORMAT("&sql_mt&",2), "
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
'response.Write(fail&" - "&int(va_f1)&" - "&va_f1&" <BR>")
fail=fail*1
if fail = 1 then

response.Redirect("altera.asp?ori=01&opt=err6&hp=err_"&url&"&obr="&obr) 

END IF



'response.Write(">>>"&periodo)




outro="U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma&""

			call GravaLog (chave,outro)

'if opt="cln" then
'outro="P:"&periodo&",D:"&co_materia&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma&", Comunicou"
'			call GravaLog (chave,outro)
'response.Redirect("comunicar.asp?or=01&nota=TB_Frequencia_Periodo&opt=ok&obr="&obr)
'else
response.Redirect("altera.asp?ori=01&opt=ok&obr="&obr)
'end if

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