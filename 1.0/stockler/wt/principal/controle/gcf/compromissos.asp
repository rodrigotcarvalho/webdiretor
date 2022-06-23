<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<script type="text/javascript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
function submitano()  
{
   var f=document.forms[0]; 
      f.submit(); 
}
function submitsistema()  
{
   var f=document.forms[1]; 
      f.submit(); 
}
function submitrapido()  
{
   var f=document.forms[2]; 
      f.submit(); 
}  
function submitforminterno()  
{
   var f=document.forms[3]; 
      f.submit(); 
	  
}
var checkflag = "false";

function check(field) {
//	alert(field.length)
	if (checkflag == "false") {
		for (i = 0; i < field.length; i++) {
			field[i].checked = true;}
			checkflag = "true";
			ativa_campo("s");			
			return "Desmarcar Todos"; 			
			}
	else {
		for (i = 0; i < field.length; i++) {
		field[i].checked = false; }
		checkflag = "false";
		ativa_campo("n");			
		return "Marcar Todos"; 	
		}
}
function validar(field){ 

for (var i = 0; i < field.length; i++)   
	if (field[i].type == "checkbox" && field[i].checked == true) {
		ativa_campo("s");
		} 
	ativa_campo("n");
}

function ativa_campo(variavel){	
	if (variavel=='s') {
		document.getElementById('acao1').disabled   = false;
		document.getElementById('acao2').disabled   = false;	
	} else {
		document.getElementById('acao1').disabled   = true;	
		document.getElementById('acao2').disabled   = true;	
	}
}

//-->
</script>
</head>
<%
nivel=4
nvg = session("chave")
chave=nvg
session("chave")=chave

ano_letivo = session("ano_letivo")
opt=Request.QueryString("opt")
pagina=Request.QueryString("pagina")
volta=Request.QueryString("v")
if (pagina=1 or pagina="1") and volta="n" then
	cod_form=request.form("busca1")
	nome_form=request.form("busca2")
	anuidade = request.form("anuidade")
	servicos = request.form("servicos")	
	aberto = request.form("aberto")
	liq_escola = request.form("liq_escola")
	liq_banco = request.form("liq_banco")
	compromissos=request.Form("compromissos")
	dia_de= request.form("dia_de")
	mes_de= request.form("mes_de")
	ano_de=request.Form("ano_de")
	dia_ate=request.Form("dia_ate")
	mes_ate=request.Form("mes_ate")
	ano_ate=request.Form("ano_ate")
	unidade=request.Form("unidade")
	curso=request.Form("curso")
	etapa=request.Form("etapa")
	turma=request.Form("turma")
else
	cod_form=Session("cod_form")
	nome_form=Session("nome_form")
	compromissos=Session("compromissos")
	anuidade = Session("anuidade")
	servicos = Session("servicos")		
	aberto = Session("aberto")
	liq_escola = Session("liq_escola")
	liq_banco = Session("liq_banco")	
	dia_de= Session("dia_de")
	mes_de= Session("mes_de")
	ano_de=Session("ano_de")
	dia_ate=Session("dia_ate")
	mes_ate=Session("mes_ate")
	ano_ate=Session("ano_ate")
	unidade=Session("unidade")
	curso=Session("curso")
	etapa=Session("etapa")
	turma=session("turma")
end if

Session("cod_form")=cod_form
Session("nome_form")=nome_form
Session("compromissos")=compromissos
Session("anuidade")=anuidade
Session("servicos")=servicos	
Session("aberto")=aberto
Session("liq_escola")=liq_escola
Session("liq_banco")=liq_banco
Session("dia_de")=dia_de
Session("mes_de")=mes_de
Session("ano_de")=ano_de
Session("dia_ate")=dia_ate
Session("mes_ate")=mes_ate
Session("ano_ate")=ano_ate
Session("unidade")=unidade
Session("curso")=curso
Session("etapa")=etapa
session("turma") =turma


		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CONa = Server.CreateObject("ADODB.Connection") 
		ABRIRa = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONa.Open ABRIRa
		
    	Set CON3 = Server.CreateObject("ADODB.Connection") 
		ABRIR3= "DBQ="& CAMINHO_cp & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3		
		
    	Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4= "DBQ="& CAMINHO_lq & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4				
		
call navegacao (CON,chave,nivel)
navega=Session("caminho")		



if unidade = "999990" or unidade = "" or isnull(unidade) then
	conjuncao="" 
else
	conjuncao=" and" 	
end if

if anuidade="" or isnull(anuidade) then
	anuidade = "n"
	anuidade_tx = "Não"	
else
	anuidade_tx = "Sim"	
end if

if servicos="" or isnull(servicos) then
	servicos = "n" 
	servicos_tx = "Não"		
else
	servicos_tx = "Sim"
end if


if compromissos="nulo" then
	if anuidade = "s" and servicos = "s" then
		anuidade_tx = "Sim"
		servicos_tx = "Sim"
		sql_compromissos = ""
	
	elseif anuidade = "s" and servicos = "n" then 	
		anuidade_tx = "Sim"
		servicos_tx = "Não"	
		sql_compromissos = "TP_Compromisso = 'COTA' AND"
	
	elseif anuidade = "n" and servicos = "s" then
		anuidade_tx = "Não"
		servicos_tx = "Sim"
		sql_compromissos = "TP_Compromisso <> 'COTA' AND"
	
	elseif anuidade = "n" and servicos = "n" then 	
		Response.Write("ERRO causado pela ausência de escolha entre anuidade e/ou serviços")
		Response.End()
	end if	
	nome_compromisso="Todos"
else
	Set RSc = Server.CreateObject("ADODB.Recordset")
	SQLc = "SELECT * FROM TB_Tipo_Compromisso WHERE CO_Compromisso='"&compromissos&"'"
	RSc.Open SQLc, CON3
	
	nome_compromisso=RSc("NO_Compromisso")	

	sql_compromissos="TP_Compromisso = '"&compromissos&"' and "
end if	



data_de=mes_de&"/"&dia_de&"/"&ano_de
data_ate=mes_ate&"/"&dia_ate&"/"&ano_ate

if dia_de<10 then
dia_de="0"&dia_de
end if
if mes_de<10 then
mes_de="0"&mes_de
end if
data_inicio=dia_de&"/"&mes_de&"/"&ano_de

if dia_ate<10 then
dia_ate="0"&dia_ate
end if
if mes_ate<10 then
mes_ate="0"&mes_ate
end if
data_fim=dia_ate&"/"&mes_ate&"/"&ano_ate	

if nome_form ="" or isnull(nome_form) then
	if cod_form="" or isnull(cod_form) then
		acao="a0"	
	else
		acao="a1"
	end if
else	
	acao="a2"		
end if

if acao="a0" then
	if unidade="999990" or unidade="" or isnull(unidade) then
		sql_un=""
		unidade_nome="Todas"
	else

		sql_un=" AND NU_Unidade= "&unidade
	
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="&unidade
		RS0.Open SQL0, CON0
		
		unidade_nome = RS0("NO_Unidade")
	end if
	
	if curso="999990" or curso="" or isnull(curso) then
		sql_cu=""
		curso_nome="Todos"
	else
		sql_cu=" AND CO_Curso='"&curso&"'"
		
		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Curso where CO_Curso='"&curso&"'"
		RS0c.Open SQL0c, CON0
		
		curso_nome = RS0c("NO_Curso")
	end if
	
	if etapa="999990" or etapa="" or isnull(etapa) then
		sql_et=""
		etapa_nome="Todas"
	else
		sql_et=" AND CO_Etapa='"&etapa&"'"
	
		Set RS0e = Server.CreateObject("ADODB.Recordset")
		SQL0e = "SELECT * FROM TB_Etapa where CO_Curso='"&curso&"' AND CO_Etapa='"&etapa&"'"
		RS0e.Open SQL0e, CON0
		
		etapa_nome = RS0e("NO_Etapa")
	end if
	
	if turma="999990" or turma="" or isnull(turma) then
		sql_tu=""
		turma_nome="Todas"
		turma_selecionada="n"		
	else
		sql_tu=" AND CO_Turma='"&turma&"'"
		turma_nome=turma
		turma_selecionada="s"
	end if

	Set RSm = Server.CreateObject("ADODB.Recordset")
	SQLm = "SELECT CO_Matricula as MATRIC FROM TB_Matriculas where NU_Ano = "&ano_letivo&sql_un&sql_cu&sql_et&sql_tu 
	RSm.Open SQLm, CONa		
	
	conta=0
	if RSm.EOF then
		sql_aluno=""
		sem_aluno="T"
	else
		sem_aluno="F"
		while not RSm.EOF 
			matric_aluno=RSm("MATRIC")
			if conta=0	then
				vetor_matric=matric_aluno
			else
				vetor_matric=vetor_matric&", "&matric_aluno		
			end if
		conta=conta+1	
		RSm.MOVENEXT
		WEND	

		sql_aluno = "CO_Matricula IN ("& vetor_matric &") and "		
		
	end if
		
	concatena_aluno ="u$"&unidade&"-"&curso&"-"&etapa&"-"&turma 	
else

	if acao="a2" then
		nome_cons = nome_form
		strProcura = Server.URLEncode(nome_form)
		strProcura = replace(strProcura,"+"," ")
		strProcura = replace(strProcura,"%27","´")
		strProcura = replace(strProcura,"%27","'")
		strProcura = replace(strProcura,"%C0,","À")
		strProcura = replace(strProcura,"%C1","Á")
		strProcura = replace(strProcura,"%C2","Â")
		strProcura = replace(strProcura,"%C3","Ã")
		strProcura = replace(strProcura,"%C9","É")
		strProcura = replace(strProcura,"%CA","Ê")
		strProcura = replace(strProcura,"%CD","Í")
		strProcura = replace(strProcura,"%D3","Ó")
		strProcura = replace(strProcura,"%D4","Ô")
		strProcura = replace(strProcura,"%D5","Õ")
		strProcura = replace(strProcura,"%DA","Ú")
		strProcura = replace(strProcura,"%DC","Ü")	
		strProcura = replace(strProcura,"%E1","à")
		strProcura = replace(strProcura,"%E1","á")
		strProcura = replace(strProcura,"%E2","â")
		strProcura = replace(strProcura,"%E3","ã")
		strProcura = replace(strProcura,"%E7","ç")
		strProcura = replace(strProcura,"%E9","é")
		strProcura = replace(strProcura,"%EA","ê")
		strProcura = replace(strProcura,"%ED","í")
		strProcura = replace(strProcura,"%F3","ó")
		strProcura = replace(strProcura,"F4","ô")
		strProcura = replace(strProcura,"F5","õ")
		strProcura = replace(strProcura,"%FA","ú")
		strProcura = replace(strProcura,"%FC","ü")	
	
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT a.CO_Matricula as MATRIC, m.NU_Unidade as UNI, m.CO_Curso as CUR, m.CO_Etapa as ETA, m.CO_Turma as TUR FROM TB_Alunos a, TB_Matriculas m where a.NO_Aluno like '%"& strProcura & "%' and a.CO_Matricula = m.CO_Matricula and m.NU_Ano = "&ano_letivo
		RS.Open SQL, CONa		
		
		IF RS.EOF then
			sem_aluno="N"		
		else
			sem_aluno="F"
			check_aluno=0
			cod_cons = RS("MATRIC")
			unidade_bd = RS("UNI")
			curso_bd = RS("CUR")
			etapa_bd = RS("ETA")
			turma_bd = RS("TUR")	
			
			sql_aluno = "CO_Matricula = "& cod_cons & " and "		
		end if							
	else	
		cod_cons=cod_form
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT a.NO_Aluno as NOME, m.NU_Unidade as UNI, m.CO_Curso as CUR, m.CO_Etapa as ETA, m.CO_Turma as TUR FROM TB_Alunos a, TB_Matriculas m where a.CO_Matricula = "& cod_cons & " and a.CO_Matricula = m.CO_Matricula and m.NU_Ano = "&ano_letivo		
		RS.Open SQL, CONa		
	
		check_aluno=0
		if RS.EOF then
			sem_aluno = "C"
		else
			sem_aluno="F"
			nome_cons = RS("NOME")		
			unidade_bd = RS("UNI")
			curso_bd = RS("CUR")
			etapa_bd = RS("ETA")
			turma_bd = RS("TUR")	
		end if			
	end if
	
	unidade_nome = "Todas"
	curso_nome = "Todos"
	etapa_nome = "Todas"	
	turma_nome = "Todas"
	sql_un=""
	sql_cu=""
	sql_et=""
	sql_tu=""

	sql_aluno = "CO_Matricula = "& cod_cons & " and "	
	
	concatena_aluno ="a$"&cod_cons 
end if	

if aberto = "s" then
	aberto_tx = "Sim"
else
	aberto_tx = "Não"
end if	

if liq_escola = "s" then
	liq_escola_tx = "Sim"
else
	liq_escola_tx = "Não"
end if	

if liq_banco = "s" then
	liq_banco_tx = "Sim"
else
	liq_banco_tx = "Não"
end if	


if aberto <> "s" and liq_banco <> "s" and liq_escola <> "s"  then
	Response.Write("ERRO causado pela ausência de escolha entre Em Aberto, Liquidado na Escola e/ou Liquidado na Rede Bancária")
	Response.End()	
end if	

Set RSC = Server.CreateObject ( "ADODB.RecordSet" ) 
RSC.Fields.Append "NU_Sequencial", 200, 255 
RSC.Fields.Append "DA_Vencimento", 7 
RSC.Fields.Append "NU_Ano_Letivo", 139, 10 
RSC.Fields.Append "NU_Contrato", 139, 10
RSC.Fields.Append "TP_Compromisso", 200, 255 
RSC.Fields.Append "NU_Parcela", 139, 10 
RSC.Fields.Append "NU_Cota", 139, 10
RSC.Fields.Append "CO_Matricula", 139, 10
RSC.Open 
if sem_aluno="C" or sem_aluno="N" or sem_aluno="T" then
	sem_link = "s"
	total_registros = 0
else
	Set RSCp= Server.CreateObject("ADODB.Recordset")
	SQLC = "SELECT * FROM TB_Compromissos where "&sql_aluno&sql_compromissos&" (DA_Vencimento BETWEEN #"&data_de&"# AND #"&data_ate&"#) order by NU_Ano_Letivo Desc, NU_Contrato,TP_Compromisso, NU_Parcela, NU_Cota, NU_Sequencial"
	RSCp.Open SQLC, CON3
	
	total_registros=0
	if RSCp.EOF then
		sem_link = "s"
	else
		sem_link = "n"	
		'total_registros=RSC.RecordCount
		
		while not RSCp.EOF
			num_seq = RSCp("NU_Sequencial")
			data_vencimento = RSCp("DA_Vencimento")
			ano_contrato = RSCp("NU_Ano_Letivo")
			nu_contrato = RSCp("NU_Contrato")		
			co_compromissos = RSCp("TP_Compromisso")			
			nu_parcela = RSCp("NU_Parcela")
			nu_cota = RSCp("NU_Cota")
			matricula_contrato = RSCp("CO_Matricula")			
			
			
			Set RSL= Server.CreateObject("ADODB.Recordset")
			SQLL = "SELECT * FROM TB_Lancamento_Realizado WHERE NU_Sequencial = "&num_seq
			RSL.Open SQLL, CON4		
		
			if RSL.eof then
				if aberto = "s" then
					RSC.AddNew
					RSC.Fields("NU_Sequencial").Value = num_seq
					RSC.Fields("DA_Vencimento").Value = data_vencimento
					RSC.Fields("NU_Ano_Letivo").Value = ano_contrato
					RSC.Fields("NU_Contrato").Value = nu_contrato
					RSC.Fields("TP_Compromisso").Value = co_compromissos
					RSC.Fields("NU_Parcela").Value = nu_parcela
					RSC.Fields("NU_Cota").Value = nu_cota	
					RSC.Fields("CO_Matricula").Value = matricula_contrato	
									
				end if
			else
				in_lanc_bancario=RSL("IN_Lancamento_Bancario")
				if in_lanc_bancario = TRUE then 
					if liq_banco = "s" then
						RSC.AddNew
						RSC.Fields("NU_Sequencial").Value = num_seq
						RSC.Fields("DA_Vencimento").Value = data_vencimento
						RSC.Fields("NU_Ano_Letivo").Value = ano_contrato
						RSC.Fields("NU_Contrato").Value = nu_contrato
						RSC.Fields("TP_Compromisso").Value = co_compromissos
						RSC.Fields("NU_Parcela").Value = nu_parcela
						RSC.Fields("NU_Cota").Value = nu_cota
						RSC.Fields("CO_Matricula").Value = matricula_contrato						
					end if						
				else
					if liq_escola = "s" then
						RSC.AddNew
						RSC.Fields("NU_Sequencial").Value = num_seq
						RSC.Fields("DA_Vencimento").Value = data_vencimento
						RSC.Fields("NU_Ano_Letivo").Value = ano_contrato
						RSC.Fields("NU_Contrato").Value = nu_contrato
						RSC.Fields("TP_Compromisso").Value = co_compromissos
						RSC.Fields("NU_Parcela").Value = nu_parcela
						RSC.Fields("NU_Cota").Value = nu_cota
						RSC.Fields("CO_Matricula").Value = matricula_contrato	
											
					end if						
				end if
			end if
		RSCp.MOVENEXT
		wend
	end if	
	total_registros=RSC.RecordCount		
end if


if acao<>"a0" and total_registros> 0 then
	onload = "onLoad = ""check(document.form.compromissos);"""
else
	onload = ""
end if
%>
<body <%response.Write(onload)%>>
<form name = "form" id="form" action="processa_form.asp" method="post">  
<% call cabecalho_novo (nivel)
	  %>    
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" class="tb_corpo">
  <tr> 
                    
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
	  </td>
  <%if sem_aluno = "C" then%>   
    <tr>
    <td><%call mensagens(nivel,303,1,0) %></td>
  </tr>
 <%elseif sem_aluno = "N" then%>   
    <tr>
    <td><%call mensagens(nivel,304,1,0) %></td>
  </tr>  
 <%elseif sem_aluno = "T" then%>   
    <tr>
    <td><%call mensagens(nivel,9714,1,0) %></td>
  </tr>    
 <%end if%>  
  <%if opt = "ok" then%>   
    <tr>
    <td><%call mensagens(nivel,805,2,0) %></td>
  </tr>
 <%end if%>   
  <tr>
    <td><%call mensagens(nivel,804,0,total_registros) %></td>
  </tr>
  <tr>
    <td class="tb_tit">Crit&eacute;rios da Pesquisa</td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td colspan="5"><table width="1000" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="147"  height="30" valign="middle"><div align="right"><font class="form_dado_texto"> Matr&iacute;cula:</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> </strong></font></div></td>
                <td width="68" height="30" valign="middle"><font class="form_dado_texto"><%RESPONSE.Write(cod_cons)%></font></td>
                <td width="141" height="30" valign="middle"><div align="right"><font class="form_dado_texto"> Nome: </font></div></td>
                <td width="304" height="30" valign="middle" ><font class="form_dado_texto"><%RESPONSE.Write(nome_cons)%><input name="selecao" type="hidden" value="<%response.write(concatena_aluno)%>"></font></td>
                <div id="usersList"></div>
                <td width="340" rowspan="2" valign="middle"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr class="tb_subtit">
                    <td height="30" align="center"> Parcelas de Anuidade </td>
                    <td height="30" align="center"> Parcelas de Servi&ccedil;os Adicionais </td>
                    </tr>
                  <tr class="form_dado_texto">
                    <td align="center"><%response.Write(anuidade_tx)%></td>                  
                    <td align="center"><%response.Write(servicos_tx)%></td>
                    </tr>
                  </table></td>
                </tr>
              <tr>
                <td  height="20">&nbsp;</td>
                <td width="68" height="20" valign="top">&nbsp;</td>
                <td height="20">&nbsp;</td>
                <td width="304" height="20" valign="top" >&nbsp;</td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td width="24%" class="tb_subtit"><div align="center"> Tipo de Compromisso </div></td>
            <td width="46%" class="tb_subtit"><div align="center">Per&iacute;odo de Vencimento</div></td>
            <td width="10%" align="center" class="tb_subtit"> Em Aberto </td>
            <td width="10%" class="tb_subtit"><div align="center"> Liquidados na Escola </div></td>
            <td width="10%" class="tb_subtit"><div align="center"> Liquidados na Rede Banc&aacute;ria </div></td>
          </tr>
          <tr>
            <td width="24%" align="center"><div align="center" class="form_dado_texto">
  <%response.Write(nome_compromisso)%>
              </div></td>
            <td width="46%"><div align="center"><font class="form_dado_texto"> 
              <%response.Write(data_inicio)%>
              at&eacute; 
              <%response.Write(data_fim)%>
              </font></div></td>
            <td width="10%" align="center"><div align="center" class="form_dado_texto">
  <%response.Write(aberto_tx)%>
              </div></td>
            <td width="10%" align="center"><span class="form_dado_texto">
              <%response.Write(liq_escola_tx)%>
            </span></td>
            <td width="10%" align="center"><div align="center" class="form_dado_texto">
              <%response.Write(liq_banco_tx)%>
            </div></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td width="220" class="tb_subtit"><div align="center">Unidade</div></td>
        <td width="280" class="tb_subtit"><div align="center">Curso</div></td>
        <td width="253" class="tb_subtit"><div align="center">Etapa</div></td>
        <td width="247" class="tb_subtit"><div align="center">Turma</div></td>
      </tr>
      <tr>
        <td width="220"> <div align="center"> <font class="form_dado_texto"> 
                            <%response.Write(unidade_nome)%>
                            </font> </div></td>
        <td width="280"> <div align="center"> <font class="form_dado_texto"> 
                            <%response.Write(curso_nome)%>
                            </font> </div></td>
        <td width="253"> <div align="center"> <font class="form_dado_texto"> 
                            <%response.Write(etapa_nome)%>
                            </font> </div></td>
        <td width="247"> <div align="center"> <font class="form_dado_texto"> 
                            <%response.Write(turma_nome)%>
                            </font> </div></td>
      </tr>
      <tr>
        <td colspan="5"><hr width="1000" /></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td class="tb_tit">Contratos</td>
  </tr>
  <tr>
    <td class="form_dado_texto" align="center">
  <% if sem_aluno<>"F" then%>    
  Nenhum compromisso encontrado!
  <%else%>
 <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr class="tb_subtit">
        <td width="20" align="center"><div align="center"> 
			<%if acao<>"a0" then%>
                       <input type="checkbox" name="todos" class="borda" value="" onClick="this.value=check(this.form.compromissos)" checked>           
            <%ELSE%>
                      <input type="checkbox" name="todos" class="borda" value="" onClick="this.value=check(this.form.compromissos)">
			<%end if%>                      
                    </div></td>
        <td width="100" align="center">N&uacute;mero</td>
        <td width="100" align="center">Matr&iacute;cula</td>
        <td width="320" align="center">Aluno</td>
        <td width="60" align="center">Data</td>
        <td width="160" align="center">Tipo de Compromisso</td>
        <td width="60" align="center">Parcela</td>
        <td width="60" align="center">Cota</td>
        <td width="60" align="center">Seq&uuml;&ecirc;ncia</td>
        <td width="60" align="center">Situa&ccedil;&atilde;o</td>
      </tr>
      
<%
		if RSC.EOF then			
%>      
      <tr>
              <td colspan="10" valign="top"> <div align="center"><font class="style1"> 
          <%response.Write("Não existem compromissos para os critérios informados!")%>
          </font></div></td>
      </tr>
<%else
		
	RSC.PageSize = 30
	 
	if Request.QueryString("pagina")="" then
		  intpagina = 1
		  RSC.MoveFirst
	else
		if cint(Request.QueryString("pagina"))<1 then
			intpagina = 1
		else
			if cint(Request.QueryString("pagina"))>RSC.PageCount then  
				intpagina = RSC.PageCount
			else
				intpagina = Request.QueryString("pagina")
			end if
		end if   
	 end if   

    RSC.AbsolutePage = intpagina
    intrec = 0
	
	check=2
	While intrec<RSC.PageSize and Not RSC.EoF
	
	 if check mod 2 =0 then
		cor = "tb_fundo_linha_par" 
	 else 
		cor ="tb_fundo_linha_impar"
	 end if
	
		data_vencimento = RSC("DA_Vencimento")
		dados_data_vencimento=split(data_vencimento,"/")
		dia_vencimento = dados_data_vencimento(0)
		if dia_vencimento< 10 then
			dia_vencimento = "0"&dia_vencimento
		end if	
		mes_vencimento = dados_data_vencimento(1)
		if mes_vencimento< 10 then
			mes_vencimento = "0"&mes_vencimento
		end if			
		data_vencimento = dia_vencimento&"/"&mes_vencimento&"/"&dados_data_vencimento(2)
		
		ano_contrato = RSC("NU_Ano_Letivo")
		nu_contrato = RSC("NU_Contrato")
		nu_contrato=formatnumber(nu_contrato,0)
		if nu_contrato<100000 then
			if nu_contrato<10000 then
				if nu_contrato<1000 then
					if nu_contrato<100 then
						if nu_contrato<10 then
							nu_contrato="00000"&replace(nu_contrato,".","")
						else
							nu_contrato="0000"&replace(nu_contrato,".","")				
						end if						
					else
						nu_contrato="000"&replace(nu_contrato,".","")				
					end if	
				else
					nu_contrato="00"&replace(nu_contrato,".","")				
				end if
			else
				nu_contrato="0"&nu_contrato					
			end if
		end if	 
		concatena_contrato = ano_contrato&"/"&nu_contrato	
			
		co_compromissos = RSC("TP_Compromisso")			
		nu_parcela = RSC("NU_Parcela")
		nu_cota = RSC("NU_Cota")		
		nu_sequencial = RSC("NU_Sequencial")	
		
		matricula_contrato = RSC("CO_Matricula")
		concatena_checkbox=matricula_contrato&"-"&ano_contrato&"$"&replace(nu_contrato,".","")&"$"&co_compromissos&"$"&nu_parcela&"$"&nu_cota&"$"&nu_sequencial	
		


		if isnull(co_compromissos) or co_compromissos= "" then
			no_compromisso=""
		else
			Set RSNp = Server.CreateObject("ADODB.Recordset")
			SQLcp = "SELECT * FROM TB_Tipo_Compromisso WHERE CO_Compromisso='"&co_compromissos&"'"
			RSNp.Open SQLcp, CON3
			
			no_compromisso=RSNp("NO_Compromisso")	
		end if			
			
		
		Set RSL= Server.CreateObject("ADODB.Recordset")
		SQLL = "SELECT * FROM TB_Lancamento_Realizado where NU_Sequencial = "&nu_sequencial&sql_situacao	
		RSL.Open SQLL, CON4		
			
		if RSL.EOF then
			if aberto = "s" then
				exibe_contrato="s"
				nome_situacao="Em Aberto"
			else
				exibe_contrato="n"				
			end if	
		else
			exibe_contrato="s"
			co_situacao=RSL("IN_Lancamento_Bancario")
			if co_situacao=TRUE then
				nome_situacao="Liq. Banco"					
			else
				nome_situacao="Liq. Escola"				
			end if	
		end if	
		
		Set RSA = Server.CreateObject("ADODB.Recordset")	
		SQLA = "SELECT a.NO_Aluno as NOME, m.NU_Unidade as UNI, m.CO_Curso as CUR, m.CO_Etapa as ETA, m.CO_Turma as TUR FROM TB_Alunos a, TB_Matriculas m where a.CO_Matricula = "& matricula_contrato & " and a.CO_Matricula = m.CO_Matricula and m.NU_Ano = "&ano_contrato	
		RSA.Open SQLA, CONa		
			
		if RSA.EOF then
			nome_contrato = "Não Informado"
			unidade_contrato = "N.I."
			curso_contrato = "N.I."
			etapa_contrato = "N.I."
			turma_contrato = "N.I."
		else
			nome_contrato = RSA("NOME")		
			co_unidade_contrato = RSA("UNI")
			co_curso_contrato = RSA("CUR")
			etapa_contrato = RSA("ETA")
			turma_contrato = RSA("TUR")	
			
			Set RS0 = Server.CreateObject("ADODB.Recordset")
			SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="&co_unidade_contrato
			RS0.Open SQL0, CON0
			
			unidade_contrato = RS0("NO_Abr")
	
			sql_cu="(Curso='"&curso&"' OR (Curso  is null)) AND"
			Set RS0c = Server.CreateObject("ADODB.Recordset")
			SQL0c = "SELECT * FROM TB_Curso where CO_Curso='"&co_curso_contrato&"'"
			RS0c.Open SQL0c, CON0
			
			curso_contrato = RS0c("NO_Abreviado_Curso")			
		end if	
		
		
		if exibe_contrato="s" then
%>      
      <tr>
        <td height="20" class="<%response.Write(cor)%>">
		<% 
		if nome_situacao="Em Aberto" then
			total_check = 1
			if acao<>"a0" and total_check = 1 then
		%>
				 <input name="compromissos" id="compromissos" type="checkbox" class="borda" value="<%response.Write(concatena_checkbox)%>" checked> <% else			
		%>
				 <input name="compromissos" id="compromissos" type="checkbox" class="borda" value="<%response.Write(concatena_checkbox)%>" > 
		<%	end if
			total_check = total_check +1
		end if	%>	                                 
</td>
        <td height="20" align="center" class="<%response.Write(cor)%>">
          <%response.Write(concatena_contrato)%>
          </td>
        <td height="20" align="center" class="<%response.Write(cor)%>">
          <%response.Write(matricula_contrato)%>
        </td>
        <td width="320" height="20" align="center" class="<%response.Write(cor)%>">
          <%response.Write(nome_contrato)%>
        </td>
        <td width="60" height="20" align="center" class="<%response.Write(cor)%>"><%response.Write(data_vencimento)%></td>
        <td width="160" height="20" align="center" class="<%response.Write(cor)%>">
<!--          <a href="alterar_bolsa.asp?mc=<%response.Write(matricula_contrato)%>&ac=<%response.Write(ano_contrato)%>&nc=<%response.Write(nu_contrato)%>">        -->
          <%response.Write(no_compromisso)%>
<!--          </a>-->
        </td>
        <td width="60" height="20" align="center" class="<%response.Write(cor)%>">
          <%response.Write(nu_parcela)%>
        </td>
        <td width="60" height="20" align="center" class="<%response.Write(cor)%>">
          <%response.Write(nu_cota)%>
        </td>
        <td width="60" height="20" align="center" class="<%response.Write(cor)%>"><a href="alterar_contrato.asp?mc=<%response.Write(matricula_contrato)%>&ac=<%response.Write(ano_contrato)%>&nc=<%response.Write(nu_contrato)%>&ns=<%response.Write(nu_sequencial)%>">
          <%response.Write(nu_sequencial)%>
          </a>
        </td>
        <td width="60" height="20" align="center" class="<%response.Write(cor)%>">
          <%response.Write(nome_situacao)%>
        </td>
      </tr>          
<% 	
			intrec=intrec+1
			check=check+1		
		end if
	RSC.MOVENEXT
	WEND
end if
	if (intrec<RSC.PageSize and intpagina = 1) or sem_link = "s" then
	else
%>         
      <tr>
        <td colspan="12" align="center" class="tb_tit">
		<%
        if intpagina>1 then
			%>
			<a href="compromissos.asp?pagina=<%=intpagina-1%>" class="linktres">Anterior</a> 
			<%
        end if 
		for contapagina=1 to RSC.PageCount 
			pagina=pagina*1
			IF contapagina=pagina then
				response.Write(contapagina)
			else
				%>
				<a href="compromissos.asp?pagina=<%=contapagina%>" class="linktres"><%response.Write(contapagina)%></a> 
				<%
			end if
		next
		if StrComp(intpagina,RSC.PageCount)<>0 then  
			%>
            <a href="compromissos.asp?pagina=<%=intpagina + 1%>" class="linktres">Próximo</a> 
            <%
		end if  
	end if 		
		%>      
        </td>
        </tr>
      
      <tr>
        <td colspan="12" align="center"></td>
      </tr>

    </table>   
<% END if%>        
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td colspan="4"><hr></td>
          </tr>
          <tr>
            <td width="25%"><div align="center">
              <input name="voltar" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','index.asp?nvg=<%=nvg%>');return document.MM_returnValue" value="Voltar">
            </div></td>
              <%if sem_link = "s" or sem_aluno="C" or sem_aluno="N" or sem_aluno="T" then
					ativa_botao="disabled"
				else
					ativa_botao=""
				end if

				'if (acao <>"a0" or turma_selecionada = "s") and sem_aluno="F" then
					ativa_incluir=""					
				'else
				'	ativa_incluir="disabled"				
				'end if
				%>            
          <td width="25%"><div align="center">
              <input name="acao" id="acao1" type="submit" class="borda_bot4" value="Excluir" <%response.Write(ativa_botao)%>>
            </div></td>
              <td width="25%"><div align="center">

              <input name="acao" id="acao2" type="submit" class="botao_prosseguir" value="Alterar" <%response.Write(ativa_botao)%>>
            </div></td>
            <td width="25%"><div align="center">
              <input name="acao" id="acao3" type="submit" class="botao_prosseguir" value="Incluir" <%response.Write(ativa_incluir)%>>
            </div></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
        </table></td>
      </tr>      
    </td>
  </tr>
      <tr>
        <td><img src="../../../../img/rodape.jpg" alt="" width="1000" height="40"></td>
      </tr>  
</table>
</form>
</body>
</html>
