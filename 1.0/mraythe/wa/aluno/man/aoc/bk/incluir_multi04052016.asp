<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<% 

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
ori = request.QueryString("or")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

trava=session("trava")
grupo=session("grupo")

codigo = request.form("alunos")
obr = request.form("obr")

obr_split=split(obr,"#!#")
unidade=obr_split(0)
curso=obr_split(1)
co_etapa=obr_split(2)
turma=obr_split(3)


obr=unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma
unidade=unidade*1
curso=curso*1
co_etapa=co_etapa*1
session("un_compara_aoc") = session("un_compara_aoc")*1
session("cs_compara_aoc") = session("cs_compara_aoc")*1
session("et_compara_aoc") = session("et_compara_aoc")*1

if (unidade = session("un_compara_aoc")) and (curso = session("cs_compara_aoc")) and (co_etapa=session("et_compara_aoc"))  then 	
	mantem_dados="S"
else
	mantem_dados="N"
end if

if unidade=999990 or isnull(unidade) or unidade="" then
	combo_prof="n"
	combo_disc="n"	
else
	if curso=999990 or isnull(curso) or curso="" then
		combo_prof="n"
		combo_disc="n"
	else
		teste_co_etapa = isnumeric(co_etapa)
		
		if teste_co_etapa= TRUE then	
			if co_etapa=999990 or isnull(co_etapa) or co_etapa="" then
				combo_prof="n"
				combo_disc="n"	
			else
'				if turma="999990" or isnull(turma) or turma="" then
'					combo_prof="n"
'					combo_disc="n"
'				else
					combo_prof="s"
					combo_disc="s"
'				end if			
			end if	
		else
			if co_etapa="999990" or isnull(co_etapa) or co_etapa="" then
				combo_prof="n"
				combo_disc="n"	
			else
'				if turma="999990" or isnull(turma) or turma="" then
'					combo_prof="n"
'					combo_disc="n"
'				else
					combo_prof="s"
					combo_disc="s"
'				end if			
			end if	
		end if		
'		if co_etapa=999990 or isnull(co_etapa) or co_etapa="" then
'			combo_prof="n"
'			combo_disc="n"	
'		else
'			if turma="999990" or isnull(turma) or turma="" then
'				combo_prof="n"
'				combo_disc="n"
'			else
'				combo_prof="s"
'				combo_disc="s"
'			end if			
'		end if		
	end if	
end if




		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
call navegacao (CON,chave,nivel)
navega=Session("caminho")		

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2

		Set CON3 = Server.CreateObject("ADODB.Connection") 
		ABRIR3 = "DBQ="& CAMINHO_o & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3
		
		
		Set CONp = Server.CreateObject("ADODB.Connection") 
		ABRIRp = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONp.Open ABRIRp		
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		
'
'		Set RS = Server.CreateObject("ADODB.Recordset")
'		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& codigo
'		RS.Open SQL, CON1
'		
'		
'nome_prof = RS("NO_Aluno")
'
'
'
'		Set RS = Server.CreateObject("ADODB.Recordset")
'		SQL = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& codigo
'		RS.Open SQL, CON1
'
'
'ano_aluno = RS("NU_Ano")
'rematricula = RS("DA_Rematricula")
'situacao = RS("CO_Situacao")
'encerramento= RS("DA_Encerramento")
'unidade= RS("NU_Unidade")
'curso= RS("CO_Curso")
'etapa= RS("CO_Etapa")
'turma= RS("CO_Turma")
'cham= RS("NU_Chamada")
'

%>
<html>
<head>
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../../../js/mm_menu.js"></script>
<script type="text/javascript" src="../../../../js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--


<!--
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
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_nbGroup(event, grpName) { //v6.0
  var i,img,nbArr,args=MM_nbGroup.arguments;
  if (event == "init" && args.length > 2) {
    if ((img = MM_findObj(args[2])) != null && !img.MM_init) {
      img.MM_init = true; img.MM_up = args[3]; img.MM_dn = img.src;
      if ((nbArr = document[grpName]) == null) nbArr = document[grpName] = new Array();
      nbArr[nbArr.length] = img;
      for (i=4; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
        if (!img.MM_up) img.MM_up = img.src;
        img.src = img.MM_dn = args[i+1];
        nbArr[nbArr.length] = img;
    } }
  } else if (event == "over") {
    document.MM_nbOver = nbArr = new Array();
    for (i=1; i < args.length-1; i+=3) if ((img = MM_findObj(args[i])) != null) {
      if (!img.MM_up) img.MM_up = img.src;
      img.src = (img.MM_dn && args[i+2]) ? args[i+2] : ((args[i+1])? args[i+1] : img.MM_up);
      nbArr[nbArr.length] = img;
    }
  } else if (event == "out" ) {
    for (i=0; i < document.MM_nbOver.length; i++) {
      img = document.MM_nbOver[i]; img.src = (img.MM_dn) ? img.MM_dn : img.MM_up; }
  } else if (event == "down") {
    nbArr = document[grpName];
    if (nbArr)
      for (i=0; i < nbArr.length; i++) { img=nbArr[i]; img.src = img.MM_up; img.MM_dn = 0; }
    document[grpName] = nbArr = new Array();
    for (i=2; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
      if (!img.MM_up) img.MM_up = img.src;
      img.src = img.MM_dn = (args[i+1])? args[i+1] : img.MM_up;
      nbArr[nbArr.length] = img;
  } }
}

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--

function checksubmit()
{
 if (document.busca.tp_ocor.value == "999999")
  {    alert("Por favor selecione um tipo de ocorr�ncia!")
   document.busca.tp_ocor.focus()
    return false
 }aula = document.busca.aula.value;
     if (aula.length > 3)
  {    alert("O valor do campo Aula deve possuir menos que 3 caracteres")
    document.busca.aula.focus()
    return false
  }
//    if (document.busca.observacao.value == "")
//  {    alert("Por favor digite uma observa��o!")
//    document.busca.observacao.focus()
//    return false
//  }
  return true
}

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
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
function submitfuncao()  
{
   var f=document.forms[3]; 
      f.submit(); 
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</head>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"> 
<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
            
    <td height="10" colspan="5" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
    </td>
          </tr>
	 <tr>             
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,315,0,0) %>
    </td>
			  </tr>			  			  
        <form action="bd.asp?opt=multi" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
          <tr>
      <td valign="top"> 
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
          <tr> 
            <td class="tb_tit"
>Dados Escolares</td>
          </tr>
          <tr> 
            <td height="10">
            <table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="100" height="10"> <div align="right"><font class="form_dado_texto"> 
                      Matr&iacute;culas: </font></div></td>
                  <td width="900" height="10"><font class="form_dado_texto"> 
                    <input name="cod" type="hidden" value="<%=codigo%>">
                    <%response.Write(codigo)%>
                    <input name="assunto" type="hidden" class="select_style" id="nome"  value="PED" size="75" maxlength="50">
                  </font></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td height="10" bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
         <tr> 
            <td>
            <table width="100%" border="0" cellspacing="0">
            <!--     <tr class="tb_subtit"> 
                  <td width="39" height="10"> <div align="center"> 
                      <%
'call GeraNomes("PORT",unidade,curso,etapa,CON0)
'no_unidades = session("no_unidades")
'no_grau = session("no_grau")
'no_serie = session("no_serie")
%>
                  Ano</div></td>
                  <td width="100" height="10"> <div align="center">Matr&iacute;cula</div></td>
                  <td width="129" height="10"> <div align="center">Cancelamento</div></td>
                  <td width="93" height="10"> <div align="center"> Situa&ccedil;&atilde;o</div></td>
                  <td width="106" height="10"> <div align="center">Unidade</div></td>
                  <td width="125" height="10"> <div align="center">Curso</div></td>
                  <td width="221" height="10"> <div align="center"> Etapa</div></td>
                  <td width="106" height="10"> <div align="center">Turma </div></td>
                  <td width="63" height="10"> <div align="center">Chamada</div></td>
                </tr>
                <tr class="tb_corpo"> 
                  <td width="39" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%'response.Write(ano_aluno)%>
                  </font></div></td>
                  <td width="100" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%'response.Write(rematricula)%>
                      </font></div></td>
                  <td width="129" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%'response.Write(encerramento)%>
                      </font></div></td>
                  <td width="93" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%
					
'		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
'		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
'		RSCONTST.Open SQLCONTST, CON0
'							
'				no_situacao = RSCONTST("TX_Descricao_Situacao")	
'					response.Write(no_situacao)%>
                      </font></div></td>
                  <td width="106" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%'response.Write(no_unidades)%>
                      </font></div></td>
                  <td width="125" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%'response.Write(no_grau)%>
                      </font></div></td>
                  <td width="221" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%'response.Write(no_serie)%>
                  </font></div></td>
                  <td width="106" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%'response.Write(turma)%>
                  </font></div></td>
                  <td width="63" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%'response.Write(cham)%>
                  </font></div></td>
                </tr>-->
                <tr class="tb_tit"> 
                  <td height="10" colspan="9">Ocorr&ecirc;ncia</td>
                </tr>
                <tr class="tb_subtit"> 
                  <td width="200" height="10">Ocorr&ecirc;ncia</td>
                  <td width="250" height="10">Professor:</td>
                  <td width="150" height="10"> <div align="left">Disciplina</div></td>
                  <td width="250" height="10">Data </td>
                  <td width="120" height="10"> <div align="left">Hora</div></td>
                  <td width="30" height="10">Aula</td>
                </tr>
                <tr class="tb_corpo"> 
                  <td width="200" height="10"><div align="left"> <font class="form_dado_texto"> 
                      <select name="tp_ocor" class="select_style" id="tp_ocor">
                        <option value="999999" selected>Selecione um tipo de ocorr�ncia</option>
                        <%
 
 		Set RSto = Server.CreateObject("ADODB.Recordset")
		SQLto = "SELECT * FROM TB_Tipo_Ocorrencia order by NO_Ocorrencia"
		RSto.Open SQLto, CON0



While not RSto.EOF
co_ocorrencia=RSto("CO_Ocorrencia")
no_ocorrencia=RSto("NO_Ocorrencia")
if grupo="ASE" and co_ocorrencia<>12 then
RSto.Movenext
else

			co_ocorrencia = co_ocorrencia*1
			Session("co_ocor_aoc")  = Session("co_ocor_aoc") *1
			if mantem_dados="S" and co_ocorrencia = Session("co_ocor_aoc") then
				o_selected = "selected"
			else
				o_selected = ""			
			end if	
%>
                        <option value="<%=co_ocorrencia%>"<%response.Write(o_selected)%>> 
                        <%Response.Write(no_ocorrencia)%>
                        </option>
                        <%

RSto.Movenext
end if
WEND
%>
                      </select>
                      </font></div></td>
                  <td width="250" height="10"><div align="left"><font class="form_dado_texto"> 
                    <%
 		Set RSmat = Server.CreateObject("ADODB.Recordset")
		turma = turma*1		
		if turma = 999990 then
			SQLmat1 = "SELECT distinct(CO_Materia_Principal), CO_Professor FROM TB_Da_Aula Where NU_Unidade="&unidade&" AND CO_Curso='"&curso&"' AND CO_Etapa='"&co_etapa&"' group by CO_Materia_Principal, CO_Professor"
		else
			SQLmat1 = "SELECT distinct CO_Materia_Principal, CO_Professor FROM TB_Da_Aula Where NU_Unidade="&unidade&" AND CO_Curso='"&curso&"' AND CO_Etapa='"&co_etapa&"' AND CO_Turma='"&turma&"' group by CO_Materia_Principal, CO_Professor"	
		end if			
		RSmat.Open SQLmat1, CON2

prof_check="nada"
prof_qtd=0
co_materia_check="nada"
While not RSmat.EOF
	co_materia=RSmat("CO_Materia_Principal")

 		Set RSnomat = Server.CreateObject("ADODB.Recordset")
		SQLnomat = "SELECT * FROM TB_Materia Where CO_Materia='"&co_materia&"'"
		RSnomat.Open SQLnomat, CON0
		
	no_materia=RSnomat("NO_Materia")

	prof=RSmat("CO_Professor")
	if prof_check=prof then
		count_prof=count_prof
	else
		prof_qtd=prof_qtd&"?"&prof
		prof_check=prof
		count_prof=count_prof+1
	end if
	if co_materia_check=co_materia then
		RSmat.Movenext
	else
	
		co_materia_check=co_materia 
		RSmat.Movenext
	end if
WEND

If count_prof=1 then
	Set RSpro = Server.CreateObject("ADODB.Recordset")
	SQLpro = "SELECT * FROM TB_Professor Where CO_Professor="&prof
	RSpro.Open SQLpro, CONp
	
	prof=RSpro("CO_Professor")
	no_prof=RSpro("NO_Professor")
	
	response.Write(no_prof)		
	
%></font>
                    <input name="co_prof" type="hidden" id="co_prof" value="<%response.Write(prof)%>"> 
                    <%
else

	if combo_prof="n" then
		response.Write("N�o dispon�vel")
	else	
		dados= split(prof_qtd, "?" )
%>
                    <select name="co_prof" class="select_style" id="co_prof">
                      <option value="999999" ></option>
                      <%
		For i=1 to ubound(dados)
 
			Set RSpro = Server.CreateObject("ADODB.Recordset")
			SQLpro = "SELECT distinct(CO_Professor), NO_Professor  FROM TB_Professor Where CO_Professor="&dados(i)&" group by CO_Professor,NO_Professor order by NO_Professor "
			RSpro.Open SQLpro, CONp

			prof=RSpro("CO_Professor")
			no_prof=RSpro("NO_Professor")
			prof = prof*1
			Session("prof_aoc") = Session("prof_aoc")*1
			if mantem_dados="S" and prof = Session("prof_aoc") then
				p_selected = "selected"
			else
				p_selected = ""			
			end if	
%>
                      <option value="<%=prof%>" <%response.Write(p_selected)%>> 
                      <%Response.Write(no_prof)%>
                      </option>
                      <%

		next
%>
                    </select> 
<%
	end if
end if					  
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

dia=dia*1
mes=mes*1
hora=hora*1
min=min*1


da_show=dia&"/"&mes&"/"&ano
data_altera=mes&"/"&dia&"/"&ano
hora_show=hora&":"&min
hora_grava=hora&":"&min

					  
					  %></font></div>				
                  </td>
                  <td width="150" height="10"> <div align="left"><font class="form_dado_texto">
<% 
	if combo_disc="n" then
		response.Write("N�o dispon�vel")
	else
%>	
                   
                      <select name="disciplina" class="select_style" id="select3" onSubmit="return checksubmit()">
                        <option value="999999"></option>
                        <%
 		Set RSmat = Server.CreateObject("ADODB.Recordset")
		turma = turma*1
		if turma = 999990 then
			SQLmat = "SELECT * FROM TB_Da_Aula Where NU_Unidade="&unidade&" AND CO_Curso='"&curso&"' AND CO_Etapa='"&co_etapa&"' order by CO_Materia_Principal"
		else
			SQLmat = "SELECT * FROM TB_Da_Aula Where NU_Unidade="&unidade&" AND CO_Curso='"&curso&"' AND CO_Etapa='"&co_etapa&"' AND CO_Turma='"&turma&"' order by CO_Materia_Principal"	
		end if		
		RSmat.Open SQLmat, CON2

co_materia_check="nada"
While not RSmat.EOF
co_materia=RSmat("CO_Materia_Principal")

 		Set RSnomat = Server.CreateObject("ADODB.Recordset")
		SQLnomat = "SELECT * FROM TB_Materia Where CO_Materia='"&co_materia&"'"
		RSnomat.Open SQLnomat, CON0

no_materia=RSnomat("NO_Materia")


if co_materia_check=co_materia then
RSmat.Movenext
else

			if mantem_dados="S" and co_materia = Session("co_materia_aoc") then
				d_selected = "selected"
			else
				d_selected = ""			
			end if	
'%>
                        <option value="<%=co_materia%>" <%response.Write(d_selected)%>> 
                        <%Response.Write(no_materia)%>
                        </option>
                        <%
co_materia_check=co_materia 
RSmat.Movenext
end if
WEND
%>
                      </select>
<%end if%>

                      </font></div></td>
                  <td width="250" height="10"><font class="form_dado_texto">
                    <select name="dia_de" id="dia_de" class="select_style">
							 <% 
							 For i =1 to 31
							 dia=dia*1
							 if dia=i then 
								if dia<10 then
								dia="0"&dia
								end if
							 %>
                                <option value="<%response.Write(i)%>" selected><%response.Write(dia)%></option>
							<% else
								if i<10 then
								i="0"&i
								end if
							%>
                                <option value="<%response.Write(i)%>"><%response.Write(i)%></option>
							<% end if 
							next
							%>	
                    </select>
                    /<select name="mes_de" id="mes_de" class="select_style">
								<%mes=mes*1
								if mes="1" or mes=1 then%>
                                <option value="1" selected>janeiro</option>
								<% else%>
                                <option value="1">janeiro</option>								
								<%end if
								if mes="2" or mes=2 then%>
                                <option value="2" selected>fevereiro</option>
								<% else%>
                                <option value="2">fevereiro</option>								
								<%end if
								if mes="3" or mes=3 then%>
                                <option value="3" selected>mar&ccedil;o</option>
								<% else%>
                                <option value="3">mar&ccedil;o</option>								
								<%end if
								if mes="4" or mes=4 then%>
                                <option value="4" selected>abril</option>
								<% else%>
                                <option value="4">abril</option>								
								<%end if
								if mes="5" or mes=5 then%>
                                <option value="5" selected>maio</option>
								<% else%>
                                <option value="5">maio</option>								
								<%end if
								if mes="6" or mes=6 then%>
                                <option value="6" selected>junho</option>
								<% else%>
                                <option value="6">junho</option>								
								<%end if
								if mes="7" or mes=7 then%>
                                <option value="7" selected>julho</option>
								<% else%>
                                <option value="7">julho</option>								
								<%end if%>
								<%if mes="8" or mes=8 then%>
                                <option value="8" selected>agosto</option>
								<% else%>
                                <option value="8">agosto</option>								
								<%end if
								if mes="9" or mes=9 then%>
                                <option value="9" selected>setembro</option>
								<% else%>
                                <option value="9">setembro</option>								
								<%end if
								if mes="10" or mes=10 then%>
                                <option value="10" selected>outubro</option>
								<% else%>
                                <option value="10">outubro</option>								
								<%end if
								if mes="11" or mes=11 then%>
                                <option value="11" selected>novembro</option>
								<% else%>
                                <option value="11">novembro</option>								
								<%end if
								if mes="12" or mes=12 then%>
                                <option value="12" selected>dezembro</option>
								<% else%>
                                <option value="12">dezembro</option>								
								<%end if%>
                              </select>
                    /<select name="ano_de" class="select_style" id="ano_de">
                                <%
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Ano_Letivo order by NU_Ano_Letivo"
		RS0.Open SQL0, CON
		while not RS0.EOF 
		ano_bd=RS0("NU_Ano_Letivo")
		
				ano_letivo=ano_letivo*1
		ano_bd=ano_bd*1

				if ano_letivo=ano_bd then%>
                                <option value="<%=ano_bd%>" selected><%=ano_bd%></option>
                                <%else%>
                                <option value="<%=ano_bd%>"><%=ano_bd%></option>
                                <%end if
		RS0.MOVENEXT
		WEND 		
				%>
                  </select></font></td>
                  <td width="120" height="10"> <div align="left"><font class="form_dado_texto" >
				                              <select name="hora_ate" id="hora_ate" class="select_style">
							 <% 
							 For i =0 to 23
							 hora=hora*1
							 if hora=i then 
							 if hora<10 then
hora="0"&hora
end if
%>
                                <option value="<%response.Write(i)%>" selected><%response.Write(hora)%></option>
							<% else%>
                                <option value="<%response.Write(i)%>"><%response.Write(i)%></option>
							<% end if 
							next
							%>
                              </select>
                              :  <select name="min_ate" id="min_ate" class="select_style">
							 <% 
							 For i =0 to 59
							 min=min*1
							 if min=i then 
if min<10 then
min="0"&min
end if%>
                                <option value="<%response.Write(i)%>" selected><%response.Write(min)%></option>
							<% else
							 if i<10 then
							 i="0"&i
							 end if
							 		%>
                                <option value="<%response.Write(i)%>"><%response.Write(i)%></option>
							<% end if 
							next
							%>
                  </select> </font></div></td>
                  <td width="30" height="10"><font class="form_dado_texto"> 
                    <input name="aula" type="text" class="textInput" id="aula" size="6">
                  </font></td>
                </tr>
                <tr class="tb_subtit"> 
                  <td height="10" colspan="6">Observa&ccedil;&atilde;o<font class="form_dado_texto"> 
                    <input name="data_altera" type="hidden" id="data_altera" value="<%=data_altera%>">
                    <input name="hora_altera" type="hidden" id="hora_altera" value="<%=hora_grava%>">
                    <input name="obr" type="hidden" value="<%response.write(obr)%>">
                    </font></td>
                </tr>
                <tr> 
                  <td height="10" colspan="6"><textarea name="observacao" cols="190" rows="5" wrap="VIRTUAL" id="observacao"></textarea></td>
                </tr>
              </table></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td><div align="center"> 
                <table width="1000" border="0" align="center" cellspacing="0">
                  <tr> 
                    <td height="24" colspan="3">
<hr></td>
                  </tr>
                  <tr> 
                    <td width="33%"> 
                      <div align="center"> 
                        <input name="alterar" type="submit" class="botao_cancelar" id="alterar" onClick="MM_goToURL('parent','select_alunos.asp?vt=s&amp;obr=<%response.Write(obr)%>');return document.MM_returnValue" value="Voltar">
                      </div></td>
                    <td width="34%">&nbsp;</td>
                    <td width="33%"> <div align="center">
                        <input name="Submit" type="submit" class="botao_prosseguir" value="Confirmar">
                      </div></td>
                  </tr>
                </table>
                <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> </font></div></td>
          </tr>
        </table></td>
    </tr>
</form>
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

</body>
</html>
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