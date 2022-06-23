<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 120 'valor em segundos
%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->
<!--#include file="../../../../inc/funcoes7.asp"-->
<!--#include file="../../../../../global/conta_alunos.asp"-->
<!--#include file="../../../../../global/tabelas_escolas.asp"-->
<!--#include file="../../../../../global/notas_calculos_diversos.asp"-->
<%
	obr = request.QueryString("obr")
	dados=obr
	dados_funcao=split(obr,"$!$")

	unidade = dados_funcao(0)
	curso = dados_funcao(1)
	co_etapa = dados_funcao(2)
	turma = dados_funcao(3)
	periodo = dados_funcao(4)
	acumulado = dados_funcao(5)
	qto_falta = dados_funcao(6)	
	ano_letivo = dados_funcao(7)
	larg_tabela = dados_funcao(8)
	alt_tabela = dados_funcao(9)

periodo_m1=4
periodo_m2=5
periodo_m3=6
	if acumulado="s" then
		for p=1 to periodo
			if p=1 then
				temp_num_periodo=p
				sigla_periodo=periodos(p,"sigla")
				temp_nomes_periodos=sigla_periodo
			else
				temp_num_periodo=temp_num_periodo&"#!#"&p
				sigla_periodo=periodos(p,"sigla")
				temp_nomes_periodos=temp_nomes_periodos&"#!#"&sigla_periodo
			end if
		next
		if qto_falta="s" then
			vetor_periodo=split(temp_nomes_periodos,"#!#")
			num_periodo=split(temp_num_periodo,"#!#")		
			for v=0 to ubound(vetor_periodo)
				if vetor_periodo(v)="BIM1" then	
					temp_num_periodo=1
					periodo_exibe=vetor_periodo(v)
				elseif vetor_periodo(v)="BIM2" then	
					temp_num_periodo=temp_num_periodo&"#!#2"
					periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)
				elseif vetor_periodo(v)="BIM3" then	
					temp_num_periodo=temp_num_periodo&"#!#3#!#0"
					periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)&"#!#QF1"
				elseif vetor_periodo(v)="BIM4" then	
					temp_num_periodo=temp_num_periodo&"#!#4#!#0#!#0"
					periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)&"#!#MB#!#QF2"					
				elseif vetor_periodo(v)="FINAL" then	
					temp_num_periodo=temp_num_periodo&"#!#5#!#0"
					periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)&"#!#MF"	
				elseif vetor_periodo(v)="REC" then	
					temp_num_periodo=temp_num_periodo&"#!#6"
					periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)						
				end if	
			next										
		else
			vetor_periodo=split(temp_nomes_periodos,"#!#")
			num_periodo=split(temp_num_periodo,"#!#")		
			for v=0 to ubound(vetor_periodo)
				if vetor_periodo(v)="BIM1" then	
					temp_num_periodo=1
					periodo_exibe=vetor_periodo(v)
				elseif vetor_periodo(v)="BIM2" then	
					temp_num_periodo=temp_num_periodo&"#!#2"
					periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)
				elseif vetor_periodo(v)="BIM3" then	
					temp_num_periodo=temp_num_periodo&"#!#3"
					periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)
				elseif vetor_periodo(v)="BIM4" then	
					temp_num_periodo=temp_num_periodo&"#!#4#!#0"
					periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)&"#!#MB"					
				elseif vetor_periodo(v)="FINAL" then	
					temp_num_periodo=temp_num_periodo&"#!#5#!#0"
					periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)&"#!#MF"	
				elseif vetor_periodo(v)="REC" then	
					temp_num_periodo=temp_num_periodo&"#!#6"
					periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)						
				end if	
			
			next					
		end if	
	else	
		temp_num_periodo=periodo
		sigla_periodo=periodos(periodo,"sigla")
		periodo_exibe=sigla_periodo
	end if
	num_periodo_detalhe=temp_num_periodo
	num_periodo=split(temp_num_periodo,"#!#")	
	javascript_periodo=periodo_exibe
	nom_periodo=split(periodo_exibe,"#!#")

	
	Set CON_wr = Server.CreateObject("ADODB.Connection") 
	ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_wr.Open ABRIR_wr
	
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CON_AL = Server.CreateObject("ADODB.Connection") 
	ABRIR_AL = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_AL.Open ABRIR_AL
	
	Set CONg = Server.CreateObject("ADODB.Connection") 
	ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONg.Open ABRIRg		
	
	Set CON3 = Server.CreateObject("ADODB.Connection") 
	ABRIR3 = "DBQ="& CAMINHO_o & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON3.Open ABRIR3

	cor_nota_vml="#FF0000"	
	cor_nota_azl="#0000FF"	
	cor_nota_prt="#000000"	
	cor_nota_vrd="#006600"	
	
	avaliacao = "VA_Media3"	
	
call GeraNomes("Port",unidade,curso,co_etapa,CON0)	
no_unidade	= 	session("no_unidades")
no_curso	=	session("no_grau")
no_etapa	=	session("no_serie")
	
	Set RS_nota = Server.CreateObject("ADODB.Recordset")
	CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"'"
	Set RS_nota = CONg.Execute(CONEXAO)


if RS_nota.EOF then
	response.Write("<div align=center><font size=2 face=Arial  color=#990000><b>Essa Turma não possui planilha de notas cadastrada!</b></font><br>")
	response.Write("<font size=2 face=Arial  color=#990000><a href=javascript:window.close()>voltar</a></font></div>")
else
	tb_nota = RS_nota("TP_Nota")
	if tb_nota ="TB_NOTA_A" then
		caminho_nota = CAMINHO_na
		opcao="A"
	elseif tb_nota="TB_NOTA_B" then
		caminho_nota = CAMINHO_nb
		opcao="B"		
	elseif tb_nota ="TB_NOTA_C" then
		caminho_nota = CAMINHO_nc
		opcao="C"
	elseif tb_nota ="TB_NOTA_D" then
		caminho_nota = CAMINHO_nd
		opcao="D"			
	elseif tb_nota ="TB_NOTA_E" then
		caminho_nota = CAMINHO_ne	
		opcao="E"					
	else
		response.Write("ERRO")
	end if	
end if	

	Set CON_N = Server.CreateObject("ADODB.Connection")
	ABRIR3 = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_N.Open ABRIR3
		
	Set RSapr = Server.CreateObject("ADODB.Recordset")
	SQLapr = "Select * from TB_Regras_Aprovacao WHERE CO_Curso = '"& curso &"' AND CO_Etapa='"&co_etapa&"'"
	Set RSapr = CON0.Execute(SQLapr)
	
	if RSapr.EOF then
		ntvml=0
	else
		ntazl= RSapr("NU_Valor_M1")		
		ntvml= RSapr("NU_Valor_M2")
		peso_m2_m1=RSapr("NU_Peso_Media_M2_M1")
		peso_m2_m2=RSapr("NU_Peso_Media_M2_M2")
		peso_m3_m1=RSapr("NU_Peso_Media_M3_M1")
		peso_m3_m2=RSapr("NU_Peso_Media_M3_M2")
		peso_m3_m3=RSapr("NU_Peso_Media_M3_M3")		
	end if

	vetor_materia=cod_cons

	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL_A = "Select * from TB_Matriculas WHERE NU_Ano="&ano_letivo&" AND NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"

	Set RS = CON_AL.Execute(SQL_A)
	IF RS.EOF Then
		alunos_vetor="nulo"
	else		
		co_aluno_check=0
		While Not RS.EOF
		nu_matricula = RS("CO_Matricula")
		nu_chamada = RS("NU_Chamada")		
		
			Set RSs = Server.CreateObject("ADODB.Recordset")
			SQL_s ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& nu_matricula&" and TB_Matriculas.NU_Ano="&ano_letivo
			Set RSs = CON_AL.Execute(SQL_s)
	
			situac=RSs("CO_Situacao")
			nome_aluno=RSs("NO_Aluno")		
	
			if situac<>"C" then
				nome_aluno=nome_aluno&" - Aluno Inativo"
			end if

			if co_aluno_check=0 then
				alunos_vetor=nu_matricula&"#!#"&nu_chamada&"#!#"&nome_aluno
			else
				alunos_vetor=alunos_vetor&"#$#"&nu_matricula&"#!#"&nu_chamada&"#!#"&nome_aluno
			end if
			co_aluno_check=co_aluno_check+1	
		RS.MOVENEXT
		WEND
	END IF			
	
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND IN_MAE=TRUE order by NU_Ordem_Boletim "
	RS.Open SQL, CON0
	co_materia_check=1
	IF RS.EOF Then
		vetor_materia_exibe="nulo"
	else
		while not RS.EOF
			co_mat_fil= RS("CO_Materia")		
			if co_materia_check=1 then
				vetor_materia=co_mat_fil
			else
				vetor_materia=vetor_materia&"#!#"&co_mat_fil
			end if
			co_materia_check=co_materia_check+1			
					
		RS.MOVENEXT
		wend						
	end if

	if vetor_materia_exibe="nulo" then
		Response.Write("Erro 1 - Não foram encontradas matérias para Etapa ='"& co_etapa &"' e Curso ="& curso)
	else
		vetor_materia_exibe=programa_aula(vetor_materia, unidade, curso, co_etapa, turma)
	end if
	vet_co_materia_detalhe=vetor_materia_exibe
	vet_co_materia= split(vetor_materia_exibe,"#!#")	
	co_materia_check=1	

	qtd_colunas=ubound(vet_co_materia)+1
	larg_min_cols=40
	width_tabela=larg_tabela-30
	height_tabela=alt_tabela-150
	height_tela=alt_tabela-30
	width_lupa=12
	width_lupa_div=width_tabela
	height_lupa_div=height_tela-24
	width_nome=230-width_lupa
	width_cabec_nome=width_nome+width_lupa	
	width_nu_chamada=20
	width_scroll=20
	width_periodo=30
	class_tit="tb_tit"
	class_subtit="tb_subtit"	
	width_tb_dados_turma=width_nu_chamada+width_nome+width_periodo+width_lupa-50
	width_else=(width_tabela-width_nome-width_nu_chamada-width_periodo-width_lupa)/qtd_colunas


	if width_else<larg_min_cols then
		width_else=larg_min_cols
		width_nome=width_nome-30
		width_tabela=width_nome+width_nu_chamada+width_periodo+(width_else*qtd_colunas)
	end if

	width_div_scroll=width_tabela-width_scroll
	width_tab_abas=width_tabela-width_tb_dados_turma-50-44
	width_aba=100
	width_tab_abas_diferenca=width_tab_abas-(2*width_aba)
	
dados_notas_detalhe=periodo_m1&"$!$"&periodo_m2&"$!$"&periodo_m3&"$!$"&ntazl&"$!$"&ntvml&"$!$"&999&"$!$"&peso_m2_m1&"$!$"&peso_m2_m2&"$!$"&peso_m3_m1&"$!$"&peso_m3_m2&"$!$"&peso_m3_m3		
session("dados_notas_detalhe")=dados_notas_detalhe
session("width_pub")=width_lupa_div
	%>	
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<style type="text/css">

/*body { font: normal 11px tahoma,arial,serif; }*/

table{margin: 0px;}
table,th,td{border-collapse: collapse;}
th,td{border-bottom: 1px solid #000000; padding: 0px;}
th span{display: block; padding: 3px}
td span{display: block; padding: 3px}
/*#lista table {width: <%response.Write(width_tabela)%>px;}
#lista th{color: #FFFFFF;background-color: #E92345;text-align: left}*/
#lista.tabContainer {width: <%response.Write(width_tabela)%>px;border: 1px solid #000000}
#lista .scrollContainer {width: <%response.Write(width_tabela)%>px;height: <%response.Write(height_tabela)%>px; overflow: auto;}
/*#lista .tabela-coluna0{width: 100px;}
#lista .tabela-coluna1{width: 150px;}
#lista .tabela-coluna2{width: 100px;}*/
.menu {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 9px;
	font-weight: normal;
	color: #000033;
	background-color: #F9F9F9;	
	border-left: 1px solid #000000;	
	border-right: 1px solid #000000;
	border-top: 1px solid #000000;
	border-bottom: 1px solid #000000;
/*padding: 2px;*/
	cursor: hand;
}
 
.menu-sel {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 9px;
	font-weight: bold;
	color: #000033;
	background-color: #FFFFFF;
	border-left: 1px solid #000000;	
	border-right: 1px solid #000000;
	border-top: 1px solid #000000;
	border-bottom: 1px solid #FFFFFF;	
	/*padding: 2px;*/
	cursor: hand;
}
 
.tb-conteudo {
	border-right: 1px solid #000000;
	border-bottom: 1px solid #000000;
}
 
.conteudo {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 9px;
	font-weight: normal;
	color: #000033;
	background-color: #FFFFFF;	
/*padding: 2px;*/
	width: <%response.Write(width_tab_abas)%>px;
	height: 40px;
}
.nome_conteudo {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 9px;
	font-weight: normal;
	color: #000033;
	background-color: #FFFFFF;
/*padding: 2px;*/
	width: <%response.Write(width_tab_abas_diferenca)%>px;
	height: 8pt;
	vertical-align: middle;
}
</style>
<script language="JavaScript" type="text/JavaScript">
<!--
<% 
parametros_funcao_jscript="celula"
for b=1 to ubound(nom_periodo) 
	parametros_funcao_jscript=parametros_funcao_jscript&",celulap"&b
next
%>
function mudar_cor_focus(<%response.Write(parametros_funcao_jscript)%>){
   celula.style.backgroundColor="#D8FF9D";
<%
for pr=0 to ubound(nom_periodo)   
	if pr=0 then
	elseif pr=1 then
   		response.Write("celulap1.style.backgroundColor=""#D8FF9D"";")
	elseif pr=2 then
   		response.Write("celulap2.style.backgroundColor=""#D8FF9D"";")
	elseif pr=3 then
   		response.Write("celulap3.style.backgroundColor=""#D8FF9D"";")
	elseif pr=4 then
   		response.Write("celulap4.style.backgroundColor=""#D8FF9D"";")
	elseif pr=5 then
   		response.Write("celulap5.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=6 then
   		response.Write("celulap6.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=7 then
   		response.Write("celulap7.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=8 then
   		response.Write("celulap8.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=9 then
   		response.Write("celulap9.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=10 then
   		response.Write("celulap10.style.backgroundColor=""#D8FF9D"";")			
	end if
next	
%>									 
}
function mudar_cor_blur_par(<%response.Write(parametros_funcao_jscript)%>){
   celula.style.backgroundColor="#FFFFFF";
<%
for pr=0 to ubound(nom_periodo)   
	if pr=0 then
	elseif pr=1 then
   		response.Write("celulap1.style.backgroundColor=""#FFFFFF"";")
	elseif pr=2 then
   		response.Write("celulap2.style.backgroundColor=""#FFFFFF"";")
	elseif pr=3 then
   		response.Write("celulap3.style.backgroundColor=""#FFFFFF"";")
	elseif pr=4 then
   		response.Write("celulap4.style.backgroundColor=""#FFFFFF"";")
	elseif pr=5 then
   		response.Write("celulap5.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=6 then
   		response.Write("celulap6.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=7 then
   		response.Write("celulap7.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=8 then
   		response.Write("celulap8.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=9 then
   		response.Write("celulap9.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=10 then
   		response.Write("celulap10.style.backgroundColor=""#FFFFFF"";")				
	end if
next	
%>   
} 
function mudar_cor_blur_impar(<%response.Write(parametros_funcao_jscript)%>){
   celula.style.backgroundColor="#FFFFE1";
<%
for pr=0 to ubound(nom_periodo)   
	if pr=0 then
	elseif pr=1 then
   		response.Write("celulap1.style.backgroundColor=""#FFFFE1"";")
	elseif pr=2 then
   		response.Write("celulap2.style.backgroundColor=""#FFFFE1"";")
	elseif pr=3 then
   		response.Write("celulap3.style.backgroundColor=""#FFFFE1"";")
	elseif pr=4 then
   		response.Write("celulap4.style.backgroundColor=""#FFFFE1"";")
	elseif pr=5 then
   		response.Write("celulap5.style.backgroundColor=""#FFFFE1"";")	
	elseif pr=6 then
   		response.Write("celulap6.style.backgroundColor=""#FFFFE1"";")	
	elseif pr=7 then
   		response.Write("celulap7.style.backgroundColor=""#FFFFE1"";")	
	elseif pr=8 then
   		response.Write("celulap8.style.backgroundColor=""#FFFFE1"";")	
	elseif pr=9 then
   		response.Write("celulap9.style.backgroundColor=""#FFFFE1"";")	
	elseif pr=10 then
   		response.Write("celulap10.style.backgroundColor=""#FFFFE1"";")				
	end if
next	
%>      
} 
function mudar_cor_blur_erro(<%response.Write(parametros_funcao_jscript)%>){
   celula.style.backgroundColor="#E4E4E4"
   <%
for pr=0 to ubound(nom_periodo)   
	if pr=0 then
	elseif pr=1 then
   		response.Write("celulap1.style.backgroundColor=""#E4E4E4"";")
	elseif pr=2 then
   		response.Write("celulap2.style.backgroundColor=""#E4E4E4"";")
	elseif pr=3 then
   		response.Write("celulap3.style.backgroundColor=""#E4E4E4"";")
	elseif pr=4 then
   		response.Write("celulap4.style.backgroundColor=""#E4E4E4"";")
	elseif pr=5 then
   		response.Write("celulap5.style.backgroundColor=""#E4E4E4"";")	
	elseif pr=6 then
   		response.Write("celulap6.style.backgroundColor=""#E4E4E4"";")
	elseif pr=7 then
   		response.Write("celulap7.style.backgroundColor=""#E4E4E4"";")	
	elseif pr=8 then
   		response.Write("celulap8.style.backgroundColor=""#E4E4E4"";")	
	elseif pr=9 then
   		response.Write("celulap9.style.backgroundColor=""#E4E4E4"";")	
	elseif pr=10 then
   		response.Write("celulap10.style.backgroundColor=""#E4E4E4"";")				
	end if
next	
%>  
}

//-->
</script>
<script>
function createXMLHTTP()
            {
                        try
                        {
                                   ajax = new ActiveXObject("Microsoft.XMLHTTP");
                        }
                        catch(e)
                        {
                                   try
                                   {
                                               ajax = new ActiveXObject("Msxml2.XMLHTTP");
                                               alert(ajax);
                                   }
                                   catch(ex)
                                   {
                                               try
                                               {
                                                           ajax = new XMLHttpRequest();
                                               }
                                               catch(exc)
                                               {
                                                            alert("Esse browser não tem recursos para uso do Ajax");
                                                            ajax = null;
                                               }
                                   }
                                   return ajax;
                        }
           
           
               var arrSignatures = ["MSXML2.XMLHTTP.5.0", "MSXML2.XMLHTTP.4.0",
               "MSXML2.XMLHTTP.3.0", "MSXML2.XMLHTTP",
               "Microsoft.XMLHTTP"];
               for (var i=0; i < arrSignatures.length; i++) {
                                                                          try {
                                                                                                             var oRequest = new ActiveXObject(arrSignatures[i]);
                                                                                                             return oRequest;
                                                                          } catch (oError) {
                                                                          }
                                      }
           
                                      throw new Error("MSXML is not installed on your system.");
                        }                                
						
						
						 function recuperarImgTbAluno(Num_Cham,MatricTb,Periodo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=imgtb", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divImgTb.innerHTML =resultado_c
recuperarImgAluno(Num_Cham,MatricTb,Periodo)
                                                           }
                                               }
 
                                               oHTTPRequest.send("matr_tb_pub=" + MatricTb);
                                   }
 								function recuperarImgAluno(Num_Cham,Matric,Periodo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=img", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divImg.innerHTML =resultado_c

                                                           }
                                               }
 
                                               oHTTPRequest.send("num_cham_pub=" + Num_Cham +"&matric_pub=" + Matric+"&periodo_pub=" + Periodo);
                                   }
 								function recuperarNota(larg_max,co_mtr,ano_letivo,curso,etapa,materia,caminho,opcao,periodo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=nt", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.div_avaliacoes.innerHTML =resultado_c

                                                           }
                                               }
 
                                               oHTTPRequest.send("larg_max_pub=" + larg_max +"&matric_pub=" + co_mtr +"&ano_pub=" + ano_letivo +"&c_pub=" + curso +"&e_pub=" + etapa +"&materia_pub=" + materia + "&caminho_pub="+caminho+ "&opcao_pub="+opcao+ "&outro_pub="+periodo);											   
                                   }	
 								function recuperarOcorrencia(NumMatric)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=ocr", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                    var resultado_o  = oHTTPRequest.responseText;
resultado_o = resultado_o.replace(/\+/g," ")
resultado_o = unescape(resultado_o)
document.all.div_ocorrencias.innerHTML =resultado_o

                                                           }
                                               }
 
                                               oHTTPRequest.send("matric_pub=" + NumMatric);
                                   }
								   
								function recuperarNomeAluno(NumMatric)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=nm", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                    var resultado_n  = oHTTPRequest.responseText;
resultado_n = resultado_n.replace(/\+/g," ")
resultado_n = unescape(resultado_n)
document.all.div_nome.innerHTML =resultado_n

                                                           }
                                               }
 
                                               oHTTPRequest.send("matric_pub=" + NumMatric);
                                   }								   
								   
								   
 								function Lupa(co_mtr,ano_letivo,unidade, curso, co_etapa, turma, vet_co_materia_detalhe, caminho_nota, tb_nota,acumulado,qto_falta,nom_periodo,num_periodo,prmtr_pub)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "detalhe.asp", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                    var resultado_a  = oHTTPRequest.responseText;
resultado_a = resultado_a.replace(/\+/g," ")
resultado_a = unescape(resultado_a)
document.all.lupa.innerHTML =resultado_a

                                                           }
                                               }
 
                                               oHTTPRequest.send("matric_pub=" + co_mtr +"&ano_pub=" + ano_letivo +"&u_pub=" + unidade +"&c_pub=" + curso +"&e_pub=" + co_etapa +"&t_pub=" + turma +"&materia_pub=" + vet_co_materia_detalhe + "&caminho_pub="+caminho_nota+ "&tb_nt="+tb_nota+ "&acum_pub=" + acumulado +"&qf_pub=" + qto_falta+"&nom_per_pub=" + nom_periodo+"&num_per_pub=" + num_periodo+"&prmtr_pub="+prmtr_pub);											   
                                   }	
								   

//Essa Função funciona apenas dentro de DETALHE.ASP
function recuperarNotaZoom(larg_max,co_mtr,ano_letivo,curso,etapa,materia,caminho2,opcao,periodo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=ntzoom", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
//Essa DIV está dentro de DETALHE.ASP
document.all.div_avaliacoes_zoom.innerHTML =resultado_c

                                                           }
                                               }
 
                                               oHTTPRequest.send("larg_max_pub=" + larg_max +"&matric_pub=" + co_mtr +"&ano_pub=" + ano_letivo +"&c_pub=" + curso +"&e_pub=" + etapa +"&materia_pub=" + materia + "&caminho_pub2="+caminho2+ "&opcao_pub="+opcao+ "&outro_pub="+periodo);											   
                                   }								   
function limpalupa(){
document.all.lupa.innerHTML =""
 }

</script>
<script language="JavaScript">
function stAba(menu,conteudo)
	{
		this.menu = menu;
		this.conteudo = conteudo;
	}
 
	var arAbas = new Array();
	arAbas[0] = new stAba('td_avaliacoes','div_avaliacoes');
	arAbas[1] = new stAba('td_ocorrencias','div_ocorrencias');
 
	function AlternarAbas(menu,conteudo)
	{
		for (i=0;i<arAbas.length;i++)
		{
			m = document.getElementById(arAbas[i].menu);
			m.className = 'menu';
			c = document.getElementById(arAbas[i].conteudo)
			c.style.display = 'none';
		}
		m = document.getElementById(menu)
		m.className = 'menu-sel';
		c = document.getElementById(conteudo)
		c.style.display = '';
	}
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
} 
function focar_load() 
{ 
<%
		n_alunos= split(alunos_vetor,"#$#")			
	
		aluno_on_load= split(n_alunos(0),"#!#")
		cod_cons_on_load=aluno_on_load(0)
		num_cham_aluno_on_load=aluno_on_load(1)
		nom_aluno_on_load=aluno_on_load(2)		

Response.Write("document.getElementById("""&num_cham_aluno_on_load&"c2"&""").focus()")
%>
} 
function focar(foco) 
{ 

document.getElementById(foco).focus()

} 
function centraliza(w,h){
//o 120 e o 16 se referem ao tamanho di cabeçalho do navegador e a barra de rolagem respectivamente
    x = parseInt((screen.width - w - 16)/2);
    y = parseInt((screen.height - h - 120)/2);
   //alert(x + '\n' + y);
    document.getElementById('alinha').style.left = x;
    document.getElementById('alinha').style.top = y;
	
//	alert('w '+x +' h '+ y)
}
function centraliza2(w){
//o 120 e o 16 se referem ao tamanho di cabeçalho do navegador e a barra de rolagem respectivamente
    x = parseInt((screen.width - w - 16)/2);
  //  y = parseInt((screen.height - h - 120)/2);
   //alert(x + '\n' + y);
    document.getElementById('lupa').style.left = x;
    document.getElementById('lupa').style.top = 10;
	
//	alert('w '+x +' h '+ y)
}
function MM_showHideLayers() { //v9.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) 
  with (document) if (getElementById && ((obj=getElementById(args[i]))!=null)) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
    obj.visibility=v; }
}
var timeout         = 3000;
var closetimer		= 0;

function mclose()
{
	div1 = document.getElementById("alinha");
	div2 = document.getElementById("fundo");	
	div1.style.visibility = 'hidden';
	div2.style.visibility = 'hidden';	
}


function mclosetime()
{
	closetimer = window.setTimeout(mclose, timeout);
}
</script>

</head>
<body onLoad="AlternarAbas('td_avaliacoes','div_avaliacoes');focar_load();recuperarNomeAluno(<% response.Write(cod_cons_on_load)%>);recuperarOcorrencia(<% response.Write(cod_cons_on_load)%>);recuperarNota(<% response.Write(width_tab_abas)%>,<% response.Write(cod_cons_on_load)%>,<% response.Write(ano_letivo)%>,<% response.Write(curso)%>,<% response.Write(co_etapa)%>,'<%response.Write(vet_co_materia(0))%>','<% response.Write(Server.URLEncode(caminho_nota))%>','<% response.Write(opcao)%>',<% response.Write(num_periodo(0))%>)">
<div id="fundo" style="position:absolute; left:0px; top:0px; width:100%; height:<%response.Write(height_tela)%>px; z-index:2; background-color: #000000; layer-background-color: #000000; border: 1px none #000000; visibility: hidden;" class="transparente"></div>
<div id="divImg">
   <div id="alinha" style="position:absolute; width:500px; z-index: 3; height: 536px; visibility: hidden;"> 
    <table border="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr> 
          <td height="16"> 
            <div align="right"> <span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide');focar('<%response.Write(num_cham&"c2")%>');mudar_cor_focus(<%response.Write(parametros_chamada_jscript)%>)">fechar</a>&nbsp;<a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide');focar('<%response.Write(num_cham&"c2")%>');mudar_cor_focus(<%response.Write(parametros_chamada_jscript)%>)"><img src="../../../../img/fecha.gif" width="20" height="16" border="0" align="absbottom"></a></font></span></div></td>
        </tr>
        <tr> 
          <td><div align="center" ><img src="../../../../img/fotos/aluno/<% Response.Write(cod_cons_on_load) %>.jpg" height="500"></div></td>
        </tr>
        <tr>
          <td height="20">
    <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <% Response.Write(nom_aluno_on_load) %>
              </font></div></td>
        </tr>
     </table>
  </div>
</div>
<div id="lupa"  style="position:absolute; left:0px; top:0px; width:<%response.Write(width_lupa_div)%>px; height:<%response.Write(height_lupa_div)%>px; z-index:2; background-color: #FFFFFF; layer-background-color: #000000; border: 1px none #000000; visibility: hidden;" >
</div>  
<div class="tabContainer" id="lista">    
	<table width="<%response.Write(width_tabela)%>" border="0" align="center" cellpadding="0" cellspacing="0">
	  <thead>
      <tr>
  		  <td width="<%response.Write(width_nu_chamada)%>" class="<%response.Write(class_tit)%>"><div align="center">N&ordm;</div></td>
          <td width="<%response.Write(width_nome)%>" class="<%response.Write(class_tit)%>">Nome</td>
          <td width="<%response.Write(width_lupa)%>" class="<%response.Write(class_tit)%>"></td>          
        <td width="<%response.Write(width_periodo)%>" class="<%response.Write(class_tit)%>"><div align="center">Per</div></td>
	<%for m=0 to ubound(vet_co_materia)%>
		  <td width="<%response.Write(width_else)%>" class="<%response.Write(class_tit)%>"><div align="center"><%response.Write(vet_co_materia(m))%></div></td>
<%	next%>  
 		  <td width="<%response.Write(width_scroll)%>" class="<%response.Write(class_tit)%>"><a href="javascript:window.close()" title="Fechar janela" ><img src="../../../../img/fecha.gif" alt="Fechar janela" width="20" height="16" border="0" align="absbottom"></a></td>      
      </thead>
	  </tr>
     </table> 
    <div class="scrollContainer" id="divscroll">   
	<table width="<%response.Write(width_div_scroll)%>" border="0" align="left" cellpadding="0" cellspacing="0">  
    <tbody>  
	<%
	check=0	
	
		for na=0 to ubound(n_alunos)
			aluno= split(n_alunos(na),"#!#")
			cod_cons=aluno(0)
			num_cham=aluno(1)
			nome_aluno=aluno(2)		
			
			if na=0 then
				cod_cons_load=cod_cons
				nome_aluno_load=nome_aluno
			end if			
		
			if right(nome_aluno,16)=" - Aluno Inativo" then
				cor = "tb_fundo_linha_falta" 
				cor2 = "tb_fundo_linha_falta" 
				onblur="mudar_cor_blur_erro"	
			else
				if check mod 2 =0 then
					cor = "tb_fundo_linha_par" 
					cor2 = "tb_fundo_linha_impar" 
					onblur="mudar_cor_blur_par"				
				else 
					cor ="tb_fundo_linha_impar"
					cor2 = "tb_fundo_linha_par" 
					onblur="mudar_cor_blur_impar"
				end if
			end if

		if ubound(nom_periodo)=0 then
			rowspan=""
		else
			rowspan="rowspan="""&ubound(nom_periodo)+1&""""
		end if	

		parametros_chamada_jscript="celula"&num_cham
		for b=1 to ubound(nom_periodo) 
			parametros_chamada_jscript=parametros_chamada_jscript&",celula"&num_cham&"p"&b
		next
		
		%>

		  <tr class="<%response.Write(cor)%>" id="<%response.Write("celula"&num_cham)%>">
			<td width="<%response.Write(width_nu_chamada)%>" <%response.Write(rowspan)%> onFocus="mudar_cor_focus(<%response.Write(parametros_chamada_jscript)%>);recuperarNomeAluno(<% response.Write(cod_cons)%>);recuperarOcorrencia(<% response.Write(cod_cons)%>);recuperarImgTbAluno(<%response.Write(num_cham)%>,<%response.Write(cod_cons)%>,'<%response.Write(javascript_periodo)%>');recuperarNota(<% response.Write(width_tab_abas)%>,<% response.Write(cod_cons)%>,<% response.Write(ano_letivo)%>,<% response.Write(curso)%>,<% response.Write(co_etapa)%>,'<%response.Write(vet_co_materia(0))%>','<% response.Write(Server.URLEncode(caminho_nota))%>','<% response.Write(opcao)%>',<% response.Write(num_periodo(0))%>)" onBlur="<%response.Write(onblur)%>(<%response.Write(parametros_chamada_jscript)%>)"><div align="center"><%response.Write(num_cham)%></div></td>
			<td width="<%response.Write(width_nome)%>" <%response.Write(rowspan)%> id="<%response.Write(num_cham)%>c2" onFocus="mudar_cor_focus(<%response.Write(parametros_chamada_jscript)%>);recuperarNomeAluno(<% response.Write(cod_cons)%>);recuperarOcorrencia(<% response.Write(cod_cons)%>);recuperarImgTbAluno(<%response.Write(num_cham)%>,<%response.Write(cod_cons)%>,'<%response.Write(javascript_periodo)%>');recuperarNota(<% response.Write(width_tab_abas)%>,<% response.Write(cod_cons)%>,<% response.Write(ano_letivo)%>,<% response.Write(curso)%>,<% response.Write(co_etapa)%>,'<%response.Write(vet_co_materia(0))%>','<% response.Write(Server.URLEncode(caminho_nota))%>','<% response.Write(opcao)%>',<% response.Write(num_periodo(0))%>)" onBlur="<%response.Write(onblur)%>(<%response.Write(parametros_chamada_jscript)%>)"><%response.Write(nome_aluno)%>
</td>
			<td width="<%response.Write(width_lupa)%>" <%response.Write(rowspan)%> onFocus="mudar_cor_focus(<%response.Write(parametros_chamada_jscript)%>);recuperarNomeAluno(<% response.Write(cod_cons)%>);recuperarOcorrencia(<% response.Write(cod_cons)%>);recuperarImgTbAluno(<%response.Write(num_cham)%>,<%response.Write(cod_cons)%>,'<%response.Write(javascript_periodo)%>');recuperarNota(<% response.Write(width_tab_abas)%>,<% response.Write(cod_cons)%>,<% response.Write(ano_letivo)%>,<% response.Write(curso)%>,<% response.Write(co_etapa)%>,'<%response.Write(vet_co_materia(0))%>','<% response.Write(Server.URLEncode(caminho_nota))%>','<% response.Write(opcao)%>',<% response.Write(num_periodo(0))%>)" onBlur="<%response.Write(onblur)%>(<%response.Write(parametros_chamada_jscript)%>)"><a href="#" title="<% response.Write(nome_aluno_load) %>" onClick="centraliza2(<% response.Write(width_tabela)%>);MM_showHideLayers('fundo','','show','lupa','','show');Lupa(<% response.Write(cod_cons)%>,<% response.Write(ano_letivo)%>,<% response.Write(unidade)%>,<% response.Write(curso)%>,'<% response.Write(co_etapa)%>','<% response.Write(turma)%>','<% response.Write(vet_co_materia_detalhe)%>','<% response.Write(Server.URLEncode(caminho_nota))%>','<% response.Write(tb_nota)%>','<% response.Write(acumulado)%>','<% response.Write(qto_falta)%>','<% response.Write(javascript_periodo)%>','<% response.Write(num_periodo_detalhe)%>','<%response.Write(parametros_chamada_jscript)%>')"> <img src="../../../../img/lupa.png" alt="Detalhar aluno" border="0" /></a></td>
          
            
			<td width="<%response.Write(width_periodo)%>" onFocus="mudar_cor_focus(<%response.Write(parametros_chamada_jscript)%>);recuperarImgTbAluno(<%response.Write(num_cham)%>,<%response.Write(cod_cons)%>,'<%response.Write(javascript_periodo)%>')" onBlur="<%response.Write(onblur)%>(<%response.Write(parametros_chamada_jscript)%>)"><div align="center"><%response.Write(nom_periodo(0))%></div></td>            
            
		<%For dsc=0 to ubound(vet_co_materia)	%>	
            <td width="<%response.Write(width_else)%>" onFocus="mudar_cor_focus(<%response.Write(parametros_chamada_jscript)%>);recuperarNomeAluno(<% response.Write(cod_cons)%>);recuperarOcorrencia(<% response.Write(cod_cons)%>);recuperarImgTbAluno(<%response.Write(num_cham)%>,<%response.Write(cod_cons)%>,'<%response.Write(javascript_periodo)%>');recuperarNota(<% response.Write(width_tab_abas)%>,<% response.Write(cod_cons)%>,<% response.Write(ano_letivo)%>,<% response.Write(curso)%>,<% response.Write(co_etapa)%>,'<%response.Write(vet_co_materia(dsc))%>','<% response.Write(Server.URLEncode(caminho_nota))%>','<% response.Write(opcao)%>',<% response.Write(num_periodo(0))%>)" onBlur="<%response.Write(onblur)%>(<%response.Write(parametros_chamada_jscript)%>)"><div align="center">
              <%
			media=ACC(unidade, curso, co_etapa, turma, cod_cons, vet_co_materia(dsc), caminho_nota, tb_nota, nom_periodo(0), num_periodo(0), periodo_m1, periodo_m2, periodo_m3, ntazl, ntvml, 999, peso_m2_m1, peso_m2_m2, peso_m3_m1, peso_m3_m2, peso_m3_m3)
			teste = isnumeric(media)			
			if teste=false then
				response.Write("&nbsp;")
			else	
				media=media*1	
				ntazl=ntazl*1
				ntvml=ntvml*1

				if media>=ntazl then	
					response.Write("<font color="&cor_nota_prt&">"&formatnumber(media,0)&"</font>")				
				elseif media>=ntvml then	
					response.Write("<font color="&cor_nota_azl&">"&formatnumber(media,0)&"</font>")
				else	
					response.Write("<font color="&cor_nota_vml&">"&formatnumber(media,0)&"</font>")	
				end if	
			end if	
			
			%>
            </div></td>
          <%NEXT%>  
		  </tr>
		<%for n=1 to ubound(nom_periodo)%>   
          <tr class="<%response.Write(cor)%>" id="<%response.Write("celula"&num_cham&"p"&n)%>">       
            <td width="<%response.Write(width_periodo)%>" onFocus="mudar_cor_focus(<%response.Write(parametros_chamada_jscript)%>);recuperarNomeAluno(<% response.Write(cod_cons)%>);recuperarOcorrencia(<% response.Write(cod_cons)%>);recuperarImgTbAluno(<%response.Write(num_cham)%>,<%response.Write(cod_cons)%>,'<%response.Write(javascript_periodo)%>')" onBlur="<%response.Write(onblur)%>(<%response.Write(parametros_chamada_jscript)%>)"><div align="center"><%response.Write(nom_periodo(n))%></div></td>
            <%For dsc=0 to ubound(vet_co_materia)	%>	
             <td width="<%response.Write(width_else)%>" onFocus="mudar_cor_focus(<%response.Write(parametros_chamada_jscript)%>);recuperarNomeAluno(<% response.Write(cod_cons)%>);recuperarOcorrencia(<% response.Write(cod_cons)%>);recuperarImgTbAluno(<%response.Write(num_cham)%>,<%response.Write(cod_cons)%>,'<%response.Write(javascript_periodo)%>');recuperarNota(<% response.Write(width_tab_abas)%>,<% response.Write(cod_cons)%>,<% response.Write(ano_letivo)%>,<% response.Write(curso)%>,<% response.Write(co_etapa)%>,'<%response.Write(vet_co_materia(dsc))%>','<% response.Write(Server.URLEncode(caminho_nota))%>','<% response.Write(opcao)%>',<% response.Write(num_periodo(n))%>)" onBlur="<%response.Write(onblur)%>(<%response.Write(parametros_chamada_jscript)%>)">
             <div align="center">
            <%
		'response.Write(vet_co_materia(dsc)&"<BR>")
		'if cod_cons=90777 and vet_co_materia(dsc)="HIST")
			media=ACC(unidade, curso, co_etapa, turma, cod_cons, vet_co_materia(dsc), caminho_nota, tb_nota, nom_periodo(n), num_periodo(n), periodo_m1, periodo_m2, periodo_m3, ntazl, ntvml, 999, peso_m2_m1, peso_m2_m2, peso_m3_m1, peso_m3_m2, peso_m3_m3)
			teste = isnumeric(media)			
			if teste=false then
				response.Write("&nbsp;")
			else	
				media=media*1	
				ntazl=ntazl*1
				ntvml=ntvml*1
				
				if (nom_periodo(n)="QF1" or nom_periodo(n)="QF2" or nom_periodo(n)="QF3") and media=0 then
					response.Write("<font color="&cor_nota_prt&"></font>")
				elseif (nom_periodo(n)="QF1" or nom_periodo(n)="QF2" or nom_periodo(n)="QF3") then
					response.Write("<font color="&cor_nota_vrd&">"&formatnumber(media,0)&"</font>")
				else	
					if media>=ntazl then	
						response.Write("<font color="&cor_nota_prt&">"&formatnumber(media,0)&"</font>")				
					elseif media>=ntvml then	
						response.Write("<font color="&cor_nota_azl&">"&formatnumber(media,0)&"</font>")
					else	
						response.Write("<font color="&cor_nota_vml&">"&formatnumber(media,0)&"</font>")	
					end if	
				end if	
			end if				
			%>
             </div>
             </td>
          <%NEXT%> 
           </tr> 
         <%next
		check=check+1
	Next
		%>
 
    </tbody> 
	</table>
     </div>      
</div><table width="<%response.Write(width_tabela)%>" border="0" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
  <tr>
    <td><hr /></td>
  </tr>
  <tr>
    <td><table width="100%" border="1"  cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="<%response.Write(width_tabela)%>" border="0" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
  <tr>
    <td width="<%response.Write(width_tb_dados_turma)%>" height="56"><table width="<%response.Write(width_tb_dados_turma)%>" border="1" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
  <tr>
    <td width="60" height="14" class="form_tit"><div align="left">UNIDADE: </div></td>
    <td width="120" height="14" class="form_dado_texto"><div align="left"><%response.Write(no_unidade)%></div></td>
  </tr> 
  <tr>
    <td width="60" height="14" class="form_tit"><div align="left">CURSO: </div></td>
    <td width="120" height="14" class="form_dado_texto"><div align="left"><%response.Write(no_curso)%></div></td>
  </tr> 
  <tr>
    <td width="60" height="14" class="form_tit"><div align="left">ETAPA: </div></td>
    <td width="120" height="14" class="form_dado_texto"><div align="left"><%response.Write(no_etapa)%></div></td>
  </tr> 
  <tr>
    <td width="60" height="14" class="form_tit"><div align="left">TURMA: </div></td>
    <td width="120" height="14" class="form_dado_texto"><div align="left"><%response.Write(turma)%></div></td>
  </tr>       
</table></td>
    <td width="50" align="center"><%

Function puxaXML() 

if left(ambiente_escola,5)="teste" then
	pasta_ambiente="wdteste"
else
	pasta_ambiente="wd"
end if
	hora = DatePart("h", now) 
	min = DatePart("n", now)
	seg= DatePart("s", now) 
  Set xmlhttp = server.CreateObject("microsoft.XMLHTTP")
  xmlhttp.open "GET","http://www.simplynet.com.br/"&pasta_ambiente&"/"&ambiente_escola&"/wa/professor/cna/acc/alunos.xml?t="&seg&min&hora,false
  xmlhttp.setrequestheader "ContentType","text/xml"
  xmlhttp.send()
  puxaXML = xmlhttp.responsexml.xml 
End Function 

dim rootElement
dim intQtdElementos

dim doc, xsldoc
  set xsldoc=server.createobject("microsoft.xmldom") 
  set doc = server.CreateObject("microsoft.xmldom")
  doc.async = false
  doc.loadxml (puxaXML())


set rootElement  = doc.documentElement
intQtdElementos = rootElement.childNodes.length
vetor_arquivos = Session("GuardaVetor")
If Not IsArray(vetor_arquivos) Then 
	vetor_arquivos = Array() 
End if

for i = 0 to intQtdElementos-1
	nome_arquivo =rootElement.childNodes(i).text
	ReDim preserve vetor_arquivos(UBound(vetor_arquivos)+1)
	vetor_arquivos(Ubound(vetor_arquivos )) = nome_arquivo
	Session("vetor_fotos") = vetor_arquivos
next

'Set Rs_ordena = Server.CreateObject ( "ADODB.RecordSet" )
'Rs_ordena.Fields.Append "nome", 201, 8000
''Vamos abrir o Recordset!
'Rs_ordena.Open
'
''Temos que percorrer agora todos os arquivos e jogar na nossa tabela virtual!
''Rs_ordena.Filter = "nome like "&cod_cons_load
'check=2
'
''if Rs_ordena.EOF then
'mostra_img="NO"
''else
'while not Rs_ordena.EOF
'nome_arquivo =Rs_ordena("nome")
'nome_jpg=cod_cons_load&".jpg"
'lowercase=lcase(nome_arquivo)
'	if nome_jpg=lowercase then
'		mostra_img="OK"
'	elseif mostra_img<>"OK" then
'		mostra_img="NO"
'	end if
'Rs_ordena.MOVENEXT
'WEND

for i = 0 to ubound(vetor_arquivos)
nome_arquivo =vetor_arquivos(i)
nome_jpg=cod_cons_load&".jpg"
lowercase=lcase(nome_arquivo)
'response.Write(nome_jpg&"-"&lowercase&"-"&mostra_img)
	if nome_jpg=lowercase then
		mostra_img="OK"
	elseif mostra_img<>"OK" then
		mostra_img="NO"
	end if
next
%><div id="divImgTb"><%
if mostra_img="OK" then
%><a href="#" title="<% response.Write(nome_aluno_load) %>" onClick="centraliza(500,536);MM_showHideLayers('fundo','','show','alinha','','show');mclosetime()"><img src="../../../../img/fotos/aluno/<% response.Write(cod_cons_load) %>.jpg" alt="<% response.Write(nome_aluno_load) %>" width="50" height="60" border="0"></a><% end if%></div></td>
    
    <td>

<table width="<%response.Write(width_tab_abas)%>" height="60" cellspacing="0" cellpadding="0"
border="0" style="border-left: 1px solid #000000;">
	<tr>
		<td height="8" width="<%response.Write(width_aba)%>" class="menu" id="td_avaliacoes"
		onClick="AlternarAbas('td_avaliacoes','div_avaliacoes')" align="center">
			Avaliações
		</td>
		<td height="8" width="<%response.Write(width_aba)%>" class="menu" id="td_ocorrencias"
		onClick="AlternarAbas('td_ocorrencias','div_ocorrencias')" align="center">
			Ocorrências
		</td>
		<td width="<%response.Write(width_tab_abas_diferenca)%>" style="border-bottom: 1px solid #000000">
			<div id="div_nome" class="nome_conteudo">


			</div>
		<td>
	</tr>
	<tr>
		<td width="<%response.Write(width_tab_abas)%>" class="tb-conteudo" colspan="3">
			<div id="div_avaliacoes" class="conteudo" style="display: none">

			</div>
			<div id="div_ocorrencias" class="conteudo" style="display: none">


			</div>
		</td>
        <td width="44"><a href="../../../../relatorios/swd102.asp?opt=acc&obr=<%response.Write(dados)%>"><img src="../../../../img/imprimir.gif" alt="Gera PDF" width="44" height="45" border="0" /></a></td>
	</tr>
</table>
    
    
    </td>
  </tr>
</table>
</td>
  </tr>
</table>
</td>
  </tr>
</table>

</body>
</html>
