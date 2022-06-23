<%On Error Resume Next%>
<!--#include file="../../../../../global/tabelas_escolas.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%

nivel=4


autoriza=Session("autoriza")
Session("autoriza")=autoriza
if autoriza="con" or  autoriza="in" or autoriza="ex" then
check_autoriza="con"
elseif autoriza="no" then
response.redirect("../../../../novologin.asp?opt=04")
end if
escola= session("escola")
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("chave")
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

unidade = request.Form("unidade")
curso = request.Form("curso")
co_etapa = request.Form("etapa")
turma = request.Form("turma")
periodo = request.Form("periodo")
mat_princ = request.Form("mat_prin")
avaliacao_form = request.Form("avaliacoes")

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2
		
		Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4
		
		Set CON_wr = Server.CreateObject("ADODB.Connection") 
		ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_wr.Open ABRIR_wr		
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Abreviado_Curso")

		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Da_Aula where NU_Unidade="& unidade &" AND CO_Curso='"& curso &"' AND CO_Etapa='"& co_etapa &"'"
		RS2.Open SQL2, CON2

tp_nota= RS2("TP_Nota")

if tp_nota="TB_NOTA_A" then
CAMINHO_n = CAMINHO_na
opcao="A"
elseif tp_nota="TB_NOTA_B" then
CAMINHO_n = CAMINHO_nb
opcao="B"
elseif tp_nota="TB_NOTA_C" then
CAMINHO_n = CAMINHO_nc
opcao="C"
elseif tp_nota="TB_NOTA_D" then
CAMINHO_n = CAMINHO_nd
opcao="D"
elseif tp_nota="TB_NOTA_E" then
CAMINHO_n = CAMINHO_ne
opcao="E"
elseif tp_nota="TB_NOTA_F" then
CAMINHO_n = CAMINHO_nf
opcao="F"
elseif tp_nota="TB_NOTA_V" then
CAMINHO_n = CAMINHO_nv
opcao="V"
elseif tp_nota="TB_NOTA_K" then
CAMINHO_n = CAMINHO_nk
opcao="K"
end if

		Set CON3 = Server.CreateObject("ADODB.Connection") 
		ABRIR3 = "DBQ="& CAMINHO_n & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3
	%>
<html>
<head>
<title>Web Acad&ecirc;mico</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--


function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
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
						
						
						 function recuperarCurso(uTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=c", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
document.all.divEtapa.innerHTML ="<select class=select_style></select>"
document.all.divPeriodo.innerHTML = "<select class=select_style></select>"
document.all.divAvaliacoes.innerHTML = "<select class=select_style></select>"
document.all.divDisciplina.innerHTML = "<select class=select_style></select>"
//recuperarEtapa()
                                                           }
                                               }

                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=e5", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divPeriodo.innerHTML = "<select class=select_style></select>"
document.all.divAvaliacoes.innerHTML = "<select class=select_style></select>"
document.all.divDisciplina.innerHTML = "<select class=select_style></select>"
//recuperarTurma()
                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }

						 function recuperarDisciplina(cTipo,eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=d3", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_d= oHTTPRequest.responseText;
resultado_d = resultado_d.replace(/\+/g," ")
resultado_d = unescape(resultado_d)
document.all.divDisciplina.innerHTML = resultado_d
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo + "&etapa_pub=" +eTipo);
                                   }

						 function recuperarPeriodo()
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=p2", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_p= oHTTPRequest.responseText;
resultado_p = resultado_p.replace(/\+/g," ")
resultado_p = unescape(resultado_p)
document.all.divPeriodo.innerHTML = resultado_p
																	   
                                                           }
                                               }

                                               oHTTPRequest.send();
                                   }
						 function recuperarAvaliacoes(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=av2", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_av= oHTTPRequest.responseText;
resultado_av = resultado_av.replace(/\+/g," ")
resultado_av = unescape(resultado_av)
document.all.divAvaliacoes.innerHTML = resultado_av
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }	
 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
                        </script>
</head> 
<body link="#CC9900" vlink="#CC9900" background="../../../../img/fundo.gif" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<table width="1000" height="670" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
                  <tr>                    
            <td height="10" class="tb_caminho"> <font color="#FFFF33" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%
	  response.Write(navega)

%>
              </font>
	</td>
  </tr>
  <%if check_autoriza="con" then%>
 <tr>                   
    <td height="10"> 
      <% call mensagens(4,9701,1,0)%>
    </td>
    </tr> 
<%else%>					
    <tr>                   
    <td height="10"> 
      <%	call mensagens(4,603,0,0) %>
    </td>
                </tr>
<%end if%>												  				  


          <tr> 
            <td valign="top">
<form name="inclusao" method="post" action="altera.asp">
                <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">
                  <tr class="tb_tit"> 
                    
            <td height="15" class="tb_tit">Grade de Aulas </td>
                  </tr>
                  <tr> 
                    
            <td><table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="166" class="tb_subtit"> <div align="center">UNIDADE 
                    </div></td>
                  <td width="166" class="tb_subtit"> <div align="center">CURSO 
                    </div></td>
                  <td width="166" class="tb_subtit"> <div align="center">ETAPA 
                  	</div></td>
                  <td width="166" class="tb_subtit"><div align="center">DISCIPLINA</div></td>
                  <td width="166" class="tb_subtit"> <div align="center">PER&Iacute;ODO</div></td>
                  <td width="166" class="tb_subtit"> <div align="center">AVALIA&Ccedil;&Otilde;ES</div></td>
                </tr>
                <tr> 
                  <td width="166"> <div align="center"> 
                      <select name="unidade" class="select_style" onChange="recuperarCurso(this.value)">
                        <%if opt<>"ok" then%>
                        <option value="999990" selected></option>
                        <%else%>
                        <option value="999990"></option>
                        <%end if
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0
While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")
NU_Unidade=NU_Unidade*1
unidade=unidade*1

if NU_Unidade=unidade then
%>
                        <option value="<%response.Write(NU_Unidade)%>" selected> 
                        <%response.Write(NO_Abr)%>
                        </option>
                        <%
else
%>
                        <option value="<%response.Write(NU_Unidade)%>"> 
                        <%response.Write(NO_Abr)%>
                        </option>
                        <%
end if
RS0.MOVENEXT
WEND
%>
                      </select>
                    </div></td>
                  <td width="166"> <div align="center"> 
                      <div id="divCurso"> 
                        <select name="curso" class="select_style" onChange="recuperarEtapa(this.value)">
                          <option value="999990"></option>
                          <%		

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT DISTINCT CO_Curso FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade
		RS0.Open SQL0, CON0	
		
While not RS0.EOF
CO_Curso = RS0("CO_Curso")

		Set RS0a = Server.CreateObject("ADODB.Recordset")
		SQL0a = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
		RS0a.Open SQL0a, CON0
		
NO_Curso = RS0a("NO_Abreviado_Curso")
if CO_Curso=curso then
%>
                          <option value="<%response.Write(CO_Curso)%>" selected> 
                          <%response.Write(NO_Curso)%>
                          </option>
                          <%
else	 	
%>
                          <option value="<%response.Write(CO_Curso)%>"> 
                          <%response.Write(NO_Curso)%>
                          </option>
                          <%
end if						
RS0.MOVENEXT
WEND
%>
                        </select>
                      </div>
                    </div></td>
                  <td width="166"> <div align="center"> 
                  	<div id="divEtapa"> 
                  		<select name="etapa" class="select_style" onChange="recuperarPeriodo();recuperarAvaliacoes(this.value)">
                  			<option value="999990"></option>
                  			<%		
		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"'"
		RS0b.Open SQL0b, CON0
				
While not RS0b.EOF
etapa = RS0b("CO_Etapa")

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&curso&"' AND CO_Etapa='"&etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")

if etapa=co_etapa then
%>
                  			<option value="<%response.Write(etapa)%>" selected> 
                  				<%response.Write(NO_Etapa)%>
                  				</option>
                  			<%
else		
%>
                  			<option value="<%response.Write(etapa)%>"> 
                  				<%response.Write(NO_Etapa)%>
                  				</option>
                  			<%
end if						
RS0b.MOVENEXT
WEND
%>
                  			</select>
                  		</div>
                  	</div></td>
                  <td width="166"><div align="center">
                  	<div id="divDisciplina">
                  		<select name="mat_prin" class="select_style" id="mat_prin" onChange="MM_callJS('recuperarPeriodo()')">
						<% if isnumeric(mat_princ) then
							sql_disc = ""
						%>
                  			<option value="999990" selected></option>
						<%else
							sql_disc = " AND CO_Materia = '"&mat_princ &"'"
						%>	
                  			<option value="999990"></option>	
						<%end if%>						
                  			<%
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso&"' order by NU_Ordem_Boletim"
		RS5.Open SQL5, CON0

	while not RS5.EOF
	co_mat_prin= RS5("CO_Materia")

		Set RS7b = Server.CreateObject("ADODB.Recordset")
		SQL7b = "SELECT * FROM TB_Materia where CO_Materia ='"& co_mat_prin &"'"
		RS7b.Open SQL7b, CON0		
		
		no_mat_prin= RS7b("NO_Materia")
		
		if co_mat_prin=mat_princ then
			selected_mat="Selected"
		else
			selected_mat=""			
		end if
		
%>
                  			<option value="<%response.Write(co_mat_prin)%>" <%response.Write(selected_mat)%>>
                  				<%response.Write(no_mat_prin)%>
                  				</option>
                  			<%						

	RS5.MOVENEXT		
	WEND
%>
                  			</select>
                  	</div>
                  	</div></td>
                  <td width="166"> <div align="center"> 
                  	<div id="divPeriodo"> 
                  		<select name="periodo" class="select_style" id="periodo">
                  			<option value="999990"></option>
                  			<%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")
NU_Periodo=NU_Periodo*1
periodo=periodo*1
if NU_Periodo=periodo then
%>
                  			<option value="<%=NU_Periodo%>" selected> 
                  				<%response.Write(NO_Periodo)%>
                  				</option>
                  			<%
else
%>
                  			<option value="<%=NU_Periodo%>"> 
                  				<%response.Write(NO_Periodo)%>
                  				</option>
                  			<%
end if
RS4.MOVENEXT
WEND%>
                  			</select>
                  		</div>
                  	</div></td>
                  <td width="166"> <div align="center"> 
                      <div id="divAvaliacoes"> 
<select name="avaliacoes" class="select_style" id="avaliacoes" onChange="MM_callJS('submitfuncao()')">
                <option value="999990"></option>
                  <%

	dados_tabela=verifica_dados_tabela(CAMINHO_n,opcao,outro)
	dados_separados=split(dados_tabela,"#$#")
	ln_nom_cols=dados_separados(4)
	nm_vars=dados_separados(5)
	nm_bd=dados_separados(6)
	avaliacoes_nomes=split(ln_nom_cols,"#!#")
	verifica_avaliacoes=split(nm_vars,"#!#")
	avaliacoes=split(nm_bd,"#!#")
	
for i=2 to UBOUND(avaliacoes_nomes)
	j=i-2
	'if avaliacoes(j)="CALCULADO" or verifica_avaliacoes(j)="media_teste" or verifica_avaliacoes(j)="media_prova"  then
	if avaliacoes(j)="CALCULADO" or verifica_avaliacoes(j)="media_teste" or verifica_avaliacoes(j)="rs" or verifica_avaliacoes(j)="rb"  then
	elseif avaliacoes(j)=avaliacao_form then
%>
                  <option value="<%response.Write(avaliacoes(j))%>" selected> 
                  <%response.Write(avaliacoes_nomes(i))%>
                  </option>
                  <%
	else
	%>
                  <option value="<%response.Write(avaliacoes(j))%>"> 
                  <%response.Write(avaliacoes_nomes(i))%>
                  </option>
					  <%
	end if	
	

NEXT
%>
                </select>                        
                      </div>
                    </div></td>
                </tr>
              </table></td>
                  </tr>
                  <tr bgcolor="#FFFFFF"> 
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>
                  <tr> 
                    
          <td><table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="245" class="tb_subtit"> 
                    <div align="center">DISCIPLINA</div></td>
                  <td width="76" class="tb_subtit"> <div align="center">TURMA</div></td>
                  <td width="37" class="tb_subtit"> <div align="center">N&ordm;</div></td>
                  <td width="360" class="tb_subtit"> <div align="center">ALUNO</div></td>
                </tr>
                <%

  		Set RS11 = Server.CreateObject("ADODB.Recordset")
		SQL11 = "SELECT * FROM TB_Matriculas where NU_Ano="& ano_letivo &" AND CO_Situacao='C' AND NU_Unidade="& unidade &" AND CO_Curso='"& curso &"' AND CO_Etapa='"& co_etapa &"' order by CO_Turma,NU_Chamada"
		RS11.Open SQL11, CON4
check=2	

while not RS11.EOF
 if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if						
CO_Matricula = RS11("CO_Matricula")
turma= RS11("CO_Turma")
NU_Chamada = RS11("NU_Chamada")
  
    	Set RS12 = Server.CreateObject("ADODB.Recordset")
		SQL12 = "SELECT * FROM TB_Alunos WHERE CO_Matricula = "&CO_Matricula &""
		RS12.Open SQL12, CON4
 NO_Aluno = RS12("NO_Aluno") 
 
  				
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"'"&sql_disc&" order by NU_Ordem_Boletim "
		RS5.Open SQL5, CON0

	while not RS5.EOF
		co_mat_prin= RS5("CO_Materia")
		mae= RS5("IN_MAE")
		fil= RS5("IN_FIL")
		nu_peso= RS5("NU_Peso")
		in_co= RS5("IN_CO")
		
		if mae=TRUE AND fil=true AND in_co=false AND isnull(nu_peso) then

		elseif mae=TRUE AND fil=false AND in_co=true AND isnull(nu_peso) then

		else
		
	
	  
	  
			Set RS7 = Server.CreateObject("ADODB.Recordset")
			SQL7 = "SELECT * FROM TB_Materia where CO_Materia_Principal='"& co_mat_prin &"'"
			RS7.Open SQL7, CON0
	
			Set RS8 = Server.CreateObject("ADODB.Recordset")
			SQL8 = "SELECT * FROM TB_Materia where CO_Materia='"& co_mat_prin &"'"
			RS8.Open SQL8, CON0
	
			no_mat= RS8("NO_Materia")
	
			if RS7.EOF Then						
			co_mat_fil = co_mat_prin
			else		
			co_mat_fil= RS7("CO_Materia")
			end if
		
		'response.Write("SQL10 = SELECT "&avaliacao&" as campo_check FROM "& tp_nota &" where CO_Matricula = "&CO_Matricula &" And CO_Materia='"& co_mat_fil&"'")
		  
			Set RS10 = Server.CreateObject("ADODB.Recordset")
			SQL10 = "SELECT "&avaliacao_form&" as campo_check FROM "& tp_nota &" where CO_Matricula = "&CO_Matricula &" AND NU_Periodo="&periodo&" And CO_Materia='"& co_mat_fil&"'"
			RS10.Open SQL10, CON3	

		' And "& campo_check &" is null 
			if RS10.EOF THEN
			%>
							<tr> 
							  <td width="245" class="<%=cor%>"> 
								<div align="center"><font class="form_dado_texto"> 
								  <%
			
					response.Write(no_mat)
			%>
								  </font></div></td>
							  <td width="76" class="<%=cor%>"> <div align="center"><font class="form_dado_texto"> 
								  <%response.Write(turma)%>
								  </font></font></div></td>
							  <td width="37" class="<%=cor%>"> <div align="center"> <font class="form_dado_texto"> 
								  <%response.Write(NU_Chamada)%>
								  </font> </font></div></td>
							  <td width="360" class="<%=cor%>"> <div align="center"><font class="form_dado_texto"> 
								  <%response.Write(NO_Aluno) %>
								  </font></font></div></td>
							</tr>
							<%				
			else
				while not RS10.EOF 
					IF RS10("campo_check")="" or isnull(RS10("campo_check")) or RS10("campo_check")=0 then
					%>
									<tr> 
									  <td width="245" class="<%=cor%>"> 
										<div align="center"><font class="form_dado_texto"> 
										  <%
					
							response.Write(no_mat)
					%>
										  </font></div></td>
									  <td width="76" class="<%=cor%>"> <div align="center"><font class="form_dado_texto"> 
										  <%response.Write(turma)%>
										  </font></font></div></td>
									  <td width="37" class="<%=cor%>"> <div align="center"> <font class="form_dado_texto"> 
										  <%response.Write(NU_Chamada)%>
										  </font> </font></div></td>
									  <td width="360" class="<%=cor%>"> <div align="center"><font class="form_dado_texto"> 
										  <%response.Write(NO_Aluno) %>
										  </font></font></div></td>
									</tr>
									<%
							
					end if
				RS10.MOVENEXT			
				WEND
			end if
		end if
	RS5.MOVENEXT
	WEND
	Response.Flush()		
check=check+1
RS11.MOVENEXT
WEND

call GravaLog (chave,avaliacao_form)
%>
              </table></td>
          </tr>
        </table>
</form>
</td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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