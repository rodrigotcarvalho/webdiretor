<%'On Error Resume Next%>
<!--#include file="caminhos.asp"-->
<!--#include file="funcoes.asp"--> 
<!--#include file="funcoes6.asp"--> 
<!--#include file="atualiza_planilha.asp"--> 
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

Set CONM = Server.CreateObject("ADODB.Connection") 
ABRIRM = "DBQ="& CAMINHO_nm & ";Driver={Microsoft Access Driver (*.mdb)}"
CONM.Open ABRIRM

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
			va_sim=request.form("simul_coord_"&i)
			va_bat=request.form("bat_coord_"&i)				
			va_bsi=request.form("bsi_coord_"&i)	
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
		va_sim=request.form("simul_coord_"&i)
		va_bat=request.form("bat_coord_"&i)				
		va_bsi=request.form("bsi_coord_"&i)	
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
	
	fail = calcula_medias_M(CONM, curso, etapa, co_materia, co_materia_pr, periodo, nu_matricula, va_av1, va_av2, va_av3, va_av4, va_av5, va_sim, va_bat, va_bsi, va_bon, va_rec)
	
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
response.Redirect(endereco_origem&"comunicar.asp?or=01&nota=TB_Nota_M&opt=ok&obr="&obr)
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
