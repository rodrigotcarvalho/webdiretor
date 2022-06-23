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
	if (checkflag == "false") {
		for (i = 0; i < field.length; i++) {
			field[i].checked = true;}
			checkflag = "true";
			return "Desmarcar Todos"; }
	else {
		for (i = 0; i < field.length; i++) {
		field[i].checked = false; }
		checkflag = "false";
		return "Marcar Todos"; }
}
function checksubmit()
{
 //valida campos checkbox das bebidas 
    var todos_inputs = document.getElementsByTagName('input');    
    for (var i=0; i<todos_inputs.length; i++){
        if(todos_inputs[i].id == "num_contrato"){
             if(todos_inputs[i].checked == true){
                   var ok = true;
                   break;
             }
             else
                   var ok = false;
         }
     }
 
     if (ok == false){
     alert('Selecione pelo menos um contrato');
     return false;
     } 
 
  return true
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
	ativos = request.form("ativos")
	cancelados = request.form("cancelados")	
	sem_parcelas = request.form("sem_parcelas")
	so_bolsistas = request.form("so_bolsistas")
	contrato=request.Form("contrato")
	dia_de= request.form("dia_de")
	mes_de= request.form("mes_de")
	ano_de=request.Form("ano_de")
	dia_ate=request.Form("dia_ate")
	mes_ate=request.Form("mes_ate")
	ano_ate=request.Form("ano_ate")
	bolsa=request.Form("bolsa")
	desconto_de=request.Form("desconto_de")
	desconto_ate=request.Form("desconto_ate")
	unidade=request.Form("unidade")
	curso=request.Form("curso")
	etapa=request.Form("etapa")
	turma=request.Form("turma")
else
	cod_form=Session("cod_form")
	nome_form=Session("nome_form")
	contrato=Session("contrato")
	ativos = Session("ativos")
	cancelados = Session("cancelados")		
	sem_parcelas = Session("sem_parcelas")
	so_bolsistas = Session("so_bolsistas")
	dia_de= Session("dia_de")
	mes_de= Session("mes_de")
	ano_de=Session("ano_de")
	dia_ate=Session("dia_ate")
	mes_ate=Session("mes_ate")
	ano_ate=Session("ano_ate")
	bolsa=Session("bolsa")
	desconto_de=Session("desconto_de")
	desconto_ate=session("desconto_ate")
	unidade=Session("unidade")
	curso=Session("curso")
	etapa=Session("etapa")
	turma=session("turma")
end if

Session("cod_form")=cod_form
Session("nome_form")=nome_form
Session("contrato")=contrato
Session("ativos")=ativos
Session("cancelados")=cancelados	
Session("sem_parcelas")=sem_parcelas
Session("so_bolsistas")=so_bolsistas
Session("dia_de")=dia_de
Session("mes_de")=mes_de
Session("ano_de")=ano_de
Session("dia_ate")=dia_ate
Session("mes_ate")=mes_ate
Session("ano_ate")=ano_ate
Session("bolsa")=bolsa
Session("desconto_de")=desconto_de
session("desconto_ate") =desconto_ate
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
		
		Set CON5 = Server.CreateObject("ADODB.Connection") 
		ABRIR5 = "DBQ="& CAMINHO_cr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON5.Open ABRIR5		
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")		



if unidade = "999990" or unidade = "" or isnull(unidade) then
	conjuncao="" 
else
	conjuncao=" and" 	
end if

if ativos="" or isnull(ativos) then
	ativos = "n"
end if

if cancelados="" or isnull(cancelados) then
	cancelados = "n" 
end if

if ativos = "s" and cancelados = "s" then
	ativos_tx = "Sim"
	cancelados_tx = "Sim"
	sql_situacao=""
		
elseif ativos = "s" and cancelados = "n" then 	
	ativos_tx = "Sim"
	cancelados_tx = "Não"
	if so_bolsistas = "s" then		
		sql_situacao="c.ST_Contrato= 'A' AND "
	else
		sql_situacao="ST_Contrato= 'A' AND "	
	end if	

elseif ativos = "n" and cancelados = "s" then
	ativos_tx = "Não"
	cancelados_tx = "Sim"
	if so_bolsistas = "s" then		
		sql_situacao="c.ST_Contrato= 'C' AND "
	else
		sql_situacao="ST_Contrato= 'C' AND "	
	end if	
elseif ativos = "n" and cancelados = "n" then 	
	Response.Write("ERRO causado pela ausência de situação de contrato")
	Response.End()
end if	




if sem_parcelas = "s" then
	sem_parcelas_tx = "Sim"
else
	sem_parcelas_tx = "Não"
end if	

if so_bolsistas = "s" then
	so_bolsistas_tx = "Sim"
else
	so_bolsistas_tx = "Não"
end if	

if contrato="" or isnull(contrato) then
	sql_contrato=""
	contrato_nome="Todos"
else
	split_contrato =split(contrato,"/")
	if ubound(split_contrato)=0 then
		contrato_pesquisa=split_contrato(0)
	else
		contrato_pesquisa=split_contrato(1)	
	end if	
	if so_bolsistas = "s" then		
		sql_contrato="c.NU_Contrato = "&contrato_pesquisa&" AND "
	else
		sql_contrato="NU_Contrato = "&contrato_pesquisa&" AND "	
	end if	
	contrato_nome=contrato
end if



if bolsa="nulo" then
	sql_bolsa=""
	bolsa_nome="Todas"	
else
	so_bolsistas = "s"
	so_bolsistas_tx = "Sim"
		
	sql_bolsa="(b.CO_Bolsa1= '"&bolsa&"' OR  b.CO_Bolsa2= '"&bolsa&"' OR  b.CO_Bolsa3= '"&bolsa&"') AND "
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Tipo_Bolsa where CO_Bolsa='"&bolsa&"'"
	RS.Open SQL, CON0	
	
	bolsa_nome=RS("NO_Bolsa")
end if

if desconto_de>0 or desconto_ate< 100 or so_bolsistas = "s" then
	so_bolsistas = "s"
	so_bolsistas_tx = "Sim"
	IF desconto_de=0 THEN
		min_desconto=1
	else
		min_desconto=desconto_de	
	end if	
	sql_desconto="(b.VA_Desconto1 BETWEEN "&min_desconto&" AND "&desconto_ate&" OR  b.VA_Desconto2 BETWEEN "&min_desconto&" AND "&desconto_ate&" OR  b.VA_Desconto3 BETWEEN "&min_desconto&" AND "&desconto_ate&") AND "		
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
	else
		sql_tu=" AND CO_Turma='"&turma&"'"
		turma_nome=turma
	end if

	Set RSm = Server.CreateObject("ADODB.Recordset")
	SQLm = "SELECT CO_Matricula as MATRIC FROM TB_Matriculas where CO_Situacao='C' "&sql_un&sql_cu&sql_et&sql_tu&" AND NU_Ano = "&ano_letivo
	'response.Write(SQLm&"<BR>")
	RSm.Open SQLm, CONa		
	
	conta=0
	if RSm.EOF then
		sql_aluno=""
	else
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
		if so_bolsistas = "s" then		
			'sql_aluno = "c.CO_Matricula IN ("& vetor_matric &") and "	
			sql_aluno=""
		else
			'sql_aluno = "CO_Matricula IN ("& vetor_matric &") and "	
			sql_aluno=""				
		end if			
	end if
		

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
		SQL = "SELECT a.NO_Aluno as NOME, a.CO_Matricula as MATRIC, m.NU_Unidade as UNI, m.CO_Curso as CUR, m.CO_Etapa as ETA, m.CO_Turma as TUR FROM TB_Alunos a, TB_Matriculas m where a.NO_Aluno like '%"& strProcura & "%' and a.CO_Matricula = m.CO_Matricula and m.NU_Ano = "&ano_letivo
		RS.Open SQL, CONa		
	
		check_aluno=0
		cod_cons = RS("MATRIC")
		nome_cons = RS("NOME")	
		unidade_bd = RS("UNI")
		curso_bd = RS("CUR")
		etapa_bd = RS("ETA")
		turma_bd = RS("TUR")								
	else	
		cod_cons=cod_form
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT a.NO_Aluno as NOME, m.NU_Unidade as UNI, m.CO_Curso as CUR, m.CO_Etapa as ETA, m.CO_Turma as TUR FROM TB_Alunos a, TB_Matriculas m where a.CO_Matricula = "& cod_cons & " and a.CO_Matricula = m.CO_Matricula and m.NU_Ano = "&ano_letivo		
		RS.Open SQL, CONa		
	
		check_aluno=0
		if RS.EOF then
			sem_aluno = "s"
		else
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
	if so_bolsistas = "s" then		
		sql_aluno = "c.CO_Matricula = "& cod_cons & " and "
	else
		sql_aluno = "CO_Matricula = "& cod_cons & " and "	
	end if	
end if	



		Set RSC= Server.CreateObject("ADODB.Recordset")
		if so_bolsistas = "s" then		
			SQLC = "SELECT * FROM TB_Contrato c, TB_Contrato_Bolsas b where "&sql_aluno&sql_situacao&sql_contrato&" (c.DT_Contrato BETWEEN #"&data_de&"# AND #"&data_ate&"#) AND "&sql_desconto&sql_bolsa&"c.CO_Matricula = b.CO_Matricula and c.NU_Contrato = b.NU_Contrato AND c.NU_Ano_Letivo = b.NU_Ano_Letivo order by c.NU_Ano_Letivo Desc,c.NU_Contrato"
		else
			SQLC = "SELECT * FROM TB_Contrato where "&sql_aluno&sql_situacao&sql_contrato&" (DT_Contrato BETWEEN #"&data_de&"# AND #"&data_ate&"#) order by NU_Ano_Letivo Desc,NU_Contrato"		
		end if	
		'response.Write(SQLC)
		RSC.Open SQLC, CON5, 3, 3
		
		'if RSC.EOF or sql_aluno="" then
		if RSC.EOF then		
			sem_link = "s"
			total_registros=0
		else
			sem_link = "n"	
			total_registros=RSC.RecordCount
		end if	
%>
<body>
<form name = "form" action="processa_form.asp" method="post" onSubmit="return checksubmit()">  
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
  <%if sem_aluno = "s" then%>   
    <tr>
    <td><%call mensagens(nivel,303,1,0) %></td>
  </tr>
 <%end if%>  
  <%if opt = "ok" then%>   
    <tr>
    <td><%call mensagens(nivel,802,2,0) %></td>
  </tr>
 <%end if%>   
  <tr>
    <td><%call mensagens(nivel,800,0,total_registros) %></td>
  </tr>
  <tr>
    <td class="tb_tit">Crit&eacute;rios da Pesquisa</td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td colspan="4"><table width="1000" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="147"  height="30" valign="middle"><div align="right"><font class="form_dado_texto"> Matr&iacute;cula:</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> </strong></font></div></td>
                <td width="68" height="30" valign="middle"><font class="form_dado_texto"><%RESPONSE.Write(cod_cons)%></font></td>
                <td width="141" height="30" valign="middle"><div align="right"><font class="form_dado_texto"> Nome: </font></div></td>
                <td width="304" height="30" valign="middle" ><font class="form_dado_texto"><%RESPONSE.Write(nome_cons)%></font></td>
                <div id="usersList"></div>
                <td width="340" rowspan="2" valign="middle"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr class="tb_subtit">
                    <td align="center">Somente Ativos</td>
                    <td align="center">Somente Cancelados</td>
                    <td align="center">Sem Parcelas</td>
                    <td align="center">Somente Bolsistas</td>
                  </tr>
                  <tr class="form_dado_texto">
                    <td align="center"><%response.Write(ativos_tx)%></td>                  
                    <td align="center"><%response.Write(cancelados_tx)%></td>
                    <td align="center"><%response.Write(sem_parcelas_tx)%></td>
                    <td align="center"><%response.Write(so_bolsistas_tx)%></td>
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
            <td width="15%" class="tb_subtit"><div align="center"> N&uacute;mero de Contrato</div></td>
            <td width="41%" class="tb_subtit"><div align="center">Per&iacute;odo</div></td>
            <td width="22%" align="center" class="tb_subtit">Tipo de Bolsa</td>
            <td width="22%" class="tb_subtit"><div align="center">Desconto</div></td>
          </tr>
          <tr>
            <td width="15%" align="center"><div align="center" class="form_dado_texto">
<%response.Write(contrato_nome)%>
            </div></td>
            <td width="41%"><div align="center"><font class="form_dado_texto"> 
                                  <%response.Write(data_inicio)%>
                                  at&eacute; 
                                  <%response.Write(data_fim)%>
                                  </font></div></td>
            <td width="22%" align="center"><div align="center" class="form_dado_texto">
<%response.Write(bolsa_nome)%>
            </div></td>
            <td width="22%"><div align="center" class="form_dado_texto">
<%response.Write(desconto_de)%>%
              at&eacute;
<%response.Write(desconto_ate)%>%
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
    <td>
 <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr class="tb_subtit">
        <td width="20" align="center"><div align="center"> 
                      <input type="checkbox" name="todos" class="borda" value="" onClick="this.value=check(this.form.num_contrato)">
                    </div></td>
        <td width="100" align="center">Data</td>
        <td width="100" align="center">N&uacute;mero</td>
        <td width="100" align="center">Matr&iacute;cula</td>
        <td width="420" align="center">Aluno</td>
        <td width="50" align="center">Bolsista?</td>
        <td width="50" align="center">Un</td>
        <td width="50" align="center">Curso</td>
        <td width="50" align="center">Etapa</td>
        <td width="50" align="center">Turma</td>
        <td width="50" align="center">Situa&ccedil;&atilde;o</td>
      </tr>
      
<%
		if RSC.EOF then
%>      
      <tr>
              <td colspan="11" valign="top"> <div align="center"><font class="style1"> 
          <%response.Write("Não existem contratos para os critérios informados!")%>
          </font></div></td>
      </tr>
<%else
	if cint(Request.QueryString("pagina"))<1 then
		intpagina = 1
		RSC.MoveFirst
	else
		if cint(Request.QueryString("pagina"))>RSC.PageCount then  
			intpagina = RSC.PageCount
		else
			intpagina = Request.QueryString("pagina")
		end if
	end if   
		
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
	
		data_contrato = RSC("DT_Contrato")
		dados_data_contrato=split(data_contrato,"/")
		dia_contrato = dados_data_contrato(0)
		if dia_contrato< 10 then
			dia_contrato = "0"&dia_contrato
		end if	
		mes_contrato = dados_data_contrato(1)
		if mes_contrato< 10 then
			mes_contrato = "0"&mes_contrato
		end if			
		data_contrato = dia_contrato&"/"&mes_contrato&"/"&dados_data_contrato(2)
		
		ano_contrato = RSC("NU_Ano_Letivo")
		nu_contrato = RSC("NU_Contrato")
		if nu_contrato<100000 then
			if nu_contrato<10000 then
				if nu_contrato<1000 then
					if nu_contrato<100 then
						if nu_contrato<10 then
							nu_contrato="00000"&nu_contrato							
						else
							nu_contrato="0000"&nu_contrato					
						end if						
					else
						nu_contrato="000"&nu_contrato					
					end if	
				else
					nu_contrato="00"&nu_contrato					
				end if
			else
				nu_contrato="0"&nu_contrato					
			end if
		end if	 
		concatena_contrato = ano_contrato&"/"&nu_contrato		
		
		matricula_contrato = RSC("CO_Matricula")
		concatena_checkbox=matricula_contrato&"-"&ano_contrato&"$"&nu_contrato	
		
		situac_contrato = RSC("ST_Contrato")
		
		if situac_contrato="A" then
			situac_contrato_nome="Ativo"		
		else
			situac_contrato_nome="Cancelado"			
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
		
		Set RSB= Server.CreateObject("ADODB.Recordset")
		SQLB = "SELECT * FROM TB_Contrato_Bolsas where CO_Matricula ="&matricula_contrato&" AND NU_Ano_Letivo="&ano_contrato&" AND NU_Contrato = "&nu_contrato
		RSB.Open SQLB, CON5	
		
		if RSB.EOF then	
			bolsista="Não"	
		else
			bolsista="Sim"			
		end if	
		

%>      
      <tr>
        <td>
		<% if acao<>"a0"then %>					
          	 <input name="num_contrato" id="num_contrato" type="checkbox" class="borda" value="<%response.Write(concatena_checkbox)%>" checked>
		<% else %>	  
          	 <input name="num_contrato" id="num_contrato" type="checkbox" class="borda" value="<%response.Write(concatena_checkbox)%>"> 
		<% end if %>	                                 
</td>
        <td align="center" class="<%response.Write(cor)%>">
          <%response.Write(data_contrato)%>
        </td>
        <td align="center" class="<%response.Write(cor)%>">
          <a href="alterar_contrato.asp?mc=<%response.Write(matricula_contrato)%>&ac=<%response.Write(ano_contrato)%>&nc=<%response.Write(nu_contrato)%>">
          <%response.Write(concatena_contrato)%>
          </a></td>
        <td align="center" class="<%response.Write(cor)%>">
          <%response.Write(matricula_contrato)%>
        </td>
        <td align="center" class="<%response.Write(cor)%>">
          <%response.Write(nome_contrato)%>
        </td>
        <td align="center" class="<%response.Write(cor)%>">
          <a href="alterar_bolsa.asp?mc=<%response.Write(matricula_contrato)%>&ac=<%response.Write(ano_contrato)%>&nc=<%response.Write(nu_contrato)%>">        
          <%response.Write(bolsista)%>
          </a>
        </td>
        <td align="center" class="<%response.Write(cor)%>">
          <%response.Write(unidade_contrato)%>
        </td>
        <td align="center" class="<%response.Write(cor)%>">
          <%response.Write(curso_contrato)%>
        </td>
        <td align="center" class="<%response.Write(cor)%>">
          <%response.Write(etapa_contrato)%>
        </td>
        <td align="center" class="<%response.Write(cor)%>">
          <%response.Write(turma_contrato)%>
        </td>
        <td align="center" class="<%response.Write(cor)%>">
          <%response.Write(situac_contrato_nome)%>
        </td>
      </tr>    
<% 		
		intrec=intrec+1
		check=check+1	
	RSC.MOVENEXT
	WEND
end if
	if (intrec<RSC.PageSize and intpagina = 1) or sem_link = "s" then
	else
%>         
      <tr>
        <td colspan="11" align="center" class="tb_tit">
		<%
        if intpagina>1 then
			%>
			<a href="contratos.asp?pagina=<%=intpagina-1%>" class="linktres">Anterior</a> 
			<%
        end if 
		for contapagina=1 to RSC.PageCount 
			pagina=pagina*1
			IF contapagina=pagina then
				response.Write(contapagina)
			else
				%>
				<a href="contratos.asp?pagina=<%=contapagina%>" class="linktres"><%response.Write(contapagina)%></a> 
				<%
			end if
		next
		if StrComp(intpagina,RSC.PageCount)<>0 then  
			%>
            <a href="contratos.asp?pagina=<%=intpagina + 1%>" class="linktres">Próximo</a> 
            <%
		end if  
	end if 		
		%>      
        </td>
        </tr>
      <tr>
        <td colspan="11" align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td colspan="4"><hr></td>
          </tr>
          <tr>
            <td width="25%"><div align="center">
              <input name="voltar" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','index.asp?nvg=<%=nvg%>');return document.MM_returnValue" value="Voltar">
            </div></td>
              <%if sem_link = "s" then
					ativa_botao="disabled"
				else
					ativa_botao=""
				end if
				%>            
          <td width="25%"><div align="center">
              <input name="acao" type="submit" class="borda_bot4" value="Cancelar" <%response.Write(ativa_botao)%>>
            </div></td>
              <td width="25%"><div align="center">

              <input name="acao" type="submit" class="botao_prosseguir" value="Alterar Contrato" <%response.Write(ativa_botao)%>>
            </div></td>
            <td width="25%"><div align="center">
              <input name="acao" type="submit" class="botao_prosseguir" value="Alterar Bolsa" <%response.Write(ativa_botao)%>>
            </div></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <!--<td>&nbsp;</td>
            <td>&nbsp;</td>-->
          </tr>
        </table></td>
      </tr>
   
    </table>   
    
    
    </td>
  </tr>
      <tr>
        <td><img src="../../../../img/rodape.jpg" alt="" width="1000" height="40"></td>
      </tr>  
</table>
</form>
</body>
</html>
