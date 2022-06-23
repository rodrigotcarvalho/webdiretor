<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/calculos.asp"-->
<!--#include file="../../../../inc/parametros.asp"-->
<!--#include file="../../../../../global/funcoes_diversas.asp"-->
<%
opt = request.QueryString("opt")
alunos_recebidos = REQUEST.QueryString("con")
obr = request.QueryString("obr")
faixa_origem = request.QueryString("fx")
order = request.QueryString("order")
nivel=4


autoriza=Session("autoriza")
Session("autoriza")=autoriza

permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
session("ano_letivo")=ano_letivo 
sistema_local=session("sistema_local")

chave=session("chave")
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

if opt="grafico" then
	dados= split(obr, "_" )
	unidade= dados(0)
	curso= dados(1)
	co_etapa= dados(2)
	periodo= dados(3)

else
	curso = request.Form("curso")
	unidade = request.Form("unidade")
	co_etapa = request.Form("etapa")
	periodo= request.Form("periodo")
	alunos_recebidos= request.Form("alunos_enviados")

end if


obr=unidade&"_"&curso&"_"&co_etapa&"_"&periodo&"_"&ano_letivo

tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")
tp_freq=tipo_divisao_ano(curso,co_etapa,"in_frequencia")

m_cons="VA_Media3"


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Abreviado_Curso")

		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' and CO_Curso ='"& curso &"'"  
		RS3.Open SQL3, CON0

 call navegacao (CON,chave,nivel)
navega=Session("caminho")
		
if RS3.EOF THEN
no_etapa="sem etapa"
else
no_etapa=RS3("NO_Etapa")
end if
	%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">

<script type="text/javascript" src="../../../../js/global.js"></script>
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
function checksubmit()
{
  if (document.inclusao.etapa.value == "")
  {    alert("Por favor, selecione uma etapa!")
    document.inclusao.etapa.focus()
    return false
  }
  if (document.inclusao.turma.value == "")
  {    alert("Por favor, selecione uma turma!")
    document.inclusao.turma.focus()
return false
}
  if (document.inclusao.mat_prin.value == "0")
  {    alert("Por favor, selecione uma disciplina!")
    document.inclusao.mat_prin.focus()
    return false
  }   
  if (document.inclusao.tabela.value == "")
  {    alert("Por favor, selecione uma tabela!")
    document.inclusao.tabela.focus()
    return false
  }                 	     
  return true
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
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
//recuperarEtapa()
                                                           }
                                               }

                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=e4", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divPeriodo.innerHTML = "<select class=select_style></select>"
//recuperarTurma()
                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t4", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t
document.all.divPeriodo.innerHTML = "<select class=select_style></select>"
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
function recuperarDisciplina(eTipo,co_prof)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=d2", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_d= oHTTPRequest.responseText;
resultado_d = resultado_d.replace(/\+/g," ")
resultado_d = unescape(resultado_d)
document.all.divDisc.innerHTML = resultado_d
document.all.divPeriodo.innerHTML = "<select class=select_style></select>"
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo + "&pr_pub=" +co_prof);
                                   }
function recuperarPeriodo(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=p", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_p= oHTTPRequest.responseText;
resultado_p = resultado_p.replace(/\+/g," ")
resultado_p = unescape(resultado_p)
document.all.divPeriodo.innerHTML = resultado_p
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif"leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" background="../../../../img/fundo_interno.gif" align="center" cellspacing="0" bgcolor="#FFFFFF">
  <tr>                    
            <td height="10" class="tb_caminho"> <font class="style-caminho">
              <%
	  response.Write(navega)

%>
              </font>
	</td>
  </tr>             <tr> 
                  
    <td height="10"> 
      <%
	call mensagens(nivel,18,0,0) 
%>
    </td>
                </tr>
                <tr> 
                  
    <td valign="top"> 
      <form name="inclusao" method="post" action="detalhar.asp?opt=detalhe&fx=<%response.Write(faixa_origem)%>&obr=<%response.Write(obr)%>&order=d" onSubmit="return checksubmit()">
                <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
                  <tr class="tb_tit"
> 
                    <td width="653" height="15" class="tb_tit"
>Segmento 
                    <input name="alunos_enviados" type="hidden" id="alunos_enviados" value="<%response.write(alunos_enviados)%>"></td>
                  </tr>
                  <tr> 
                    
            <td><table width="998" border="0" cellspacing="0">
                <tr> 
                  <td width="25%" class="tb_subtit"> <div align="center">UNIDADE 
                    </div></td>
                  <td width="25%" class="tb_subtit"> <div align="center">CURSO 
                    </div></td>
                  <td width="25%" class="tb_subtit"> <div align="center">ETAPA 
                    </div></td>
                  <td width="25%" class="tb_subtit"> <div align="center">PER&Iacute;ODO</div></td>
                </tr>
                <tr>
                  <td background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF"> 
                    <div align="center"> 
                      <select name="unidade" class="select_style" onChange="recuperarCurso(this.value)">
                        <%		
			Set RS0 = Server.CreateObject("ADODB.Recordset")
			SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
			RS0.Open SQL0, CON0
While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")
unidade=unidade*1
NU_Unidade=NU_Unidade*1
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
                  <td background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF"> 
                    <div align="center"> 
                      <div id="divCurso"> 
                        <select name="curso" class="select_style" onChange="recuperarEtapa(this.value)">
                          <%		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT Distinct CO_Curso FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade
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
                  <td background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF"> 
                    <div align="center"> 
                      <div id="divEtapa"> 
                        <select name="etapa" class="select_style" onChange="recuperarPeriodo(this.value)">
                          <%		

		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"'"
		RS0b.Open SQL0b, CON0
		
		
While not RS0b.EOF
Etapa = RS0b("CO_Etapa")


		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&curso&"' AND CO_Etapa='"&Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
if Etapa=co_etapa then
%>
                          <option value="<%response.Write(Etapa)%>" selected> 
                          <%response.Write(NO_Etapa)%>
                          </option>
                          <%
else
%>
                          <option value="<%response.Write(Etapa)%>"> 
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
                  <td width="25%"> <div id="divPeriodo" align="center"> 
                      <select name="periodo" class="select_style" id="periodo" onChange="MM_callJS('submitfuncao()')">
                        <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")
periodo=periodo*1
NU_Periodo=NU_Periodo*1
%>
                        <% if NU_Periodo=periodo then%>
                        <option value="<%=NU_Periodo%>" selected><%=NO_Periodo%></option>
                        <%else%>
                        <option value="<%=NU_Periodo%>"><%=NO_Periodo%></option>
                        <%end if%>
                        <%RS4.MOVENEXT
WEND%>
                      </select>
                    </div></td>
                </tr>
                <tr>
                  <td colspan="4" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF"><hr></td>
                </tr>
              </table></td>
                  </tr>
                  <tr> 
                    
            <td align="center" valign="top">
              <%


	Set RSFIL = Server.CreateObject("ADODB.Recordset")
		SQLFIL = "SELECT * FROM TB_Da_Aula where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"'" 
		RSFIL.Open SQLFIL, CON2

'response.Write("		SQLFIL = SELECT * FROM TB_Da_Aula where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"' AND CO_Materia_Principal='"&co_mat_fil&"'" )		
	notaFIL=RSFIL("TP_Nota")
co_mat_prin = RSFIL("CO_Materia")

if notaFIL ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na

elseif notaFIL="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb

elseif notaFIL ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc
		
elseif notaFIL ="TB_NOTA_D" then
		CAMINHOn = CAMINHO_nd

elseif notaFIL ="TB_NOTA_E" then
		CAMINHOn = CAMINHO_ne	
				
else
		response.Write("ERRO")
end if

	
		Set CONn = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONn.Open ABRIRn

		Set RSNN = Server.CreateObject("ADODB.Recordset")
		CONEXAONN = "Select * from TB_Programa_Aula WHERE CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa&"' order by NU_Ordem_Boletim"
		Set RSNN = CON0.Execute(CONEXAONN)	
				
nu_materia_check = 0
while not RSNN.EOF
	curso=curso*1
	if isnumeric(co_etapa) then
		co_etapa=co_etapa*1
	end if	
	co_materia= RSNN("CO_Materia")	
	in_mae= RSNN("IN_MAE")
	in_fil= RSNN("IN_FIL")	
	if (co_materia="ARTC" and curso=1 and co_etapa<6) then
	else
	
		if in_mae=TRUE AND in_fil=TRUE then
			wrk_tipo_materia="mcf"
		elseif in_mae=TRUE AND in_fil=FALSE then	
			wrk_tipo_materia="msf"	
		elseif in_mae=FALSE AND in_fil=TRUE then
			wrk_tipo_materia="f"
		else				
			response.Write("ERRO na classificação do tipo de matéria")
			response.End()
		end if		
		
		if wrk_tipo_materia<>"f" then	
			if nu_materia_check = 0 then
				vetor_materias=co_materia
				vetor_tipo_materia=wrk_tipo_materia
			else
				vetor_materias=vetor_materias&"#!#"&co_materia
				vetor_tipo_materia=vetor_tipo_materia&"#!#"&wrk_tipo_materia
			end if
		end if	
		nu_materia_check=nu_materia_check+1			
	end if	
RSNN.MOVENEXT
wend

co_mat_cons=split(vetor_materias,"#!#")
tp_mat_cons=split(vetor_tipo_materia,"#!#")

		Set RSt0 = Server.CreateObject("ADODB.Recordset")
		SQLt0 = "SELECT * FROM TB_Matriculas where NU_Ano ="& ano_letivo &" AND NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' And CO_Situacao='C' order by NU_Chamada"
		RSt0.Open SQLt0, CON4
		
nu_media_check = 1
while not RSt0.EOF
	nu_matricula = RSt0("CO_Matricula")
	acumula_aluno=0		
	media_numerador=0
	media_denominador=0
	pula_aluno="N"		
	for q=0 to ubound(co_mat_cons)
		disciplina=co_mat_cons(q)
		tp_disc=tp_mat_cons(q)		

		if tp_disc="msf" then
			'media_aluno=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, nu_matricula, disciplina, CAMINHOn, notaFIL, periodo)
			media_aluno=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, nu_matricula, disciplina, disciplina, CONn , notaFIL, periodo, "VA_Media2", outro)				
			if media_aluno="&nbsp;" then
				pula_aluno="S"				
			else
				media_denominador=media_denominador+1
			end if		

		elseif	tp_disc="mcf" then
			conta_filhas = 0
			
			Set RSNNa = Server.CreateObject("ADODB.Recordset")
			CONEXAONNa = "Select * from TB_Materia WHERE CO_Materia_Principal = '"& disciplina &"' order by NU_Ordem_Boletim"
			Set RSNNa = CON0.Execute(CONEXAONNa)	
			
			if RSNNa.EOF then
				Response.Write("Erro ao localizar mat&eacute;rias filhas para "&disciplina&" em TB_Materia")
				Response.end()
			else	
				while not RSNNa.EOF
				
					filha=RSNNa("CO_Materia")
	
					if conta_filhas = 0 then
						vetor_filhas=filha
					else
						vetor_filhas=vetor_filhas&"#!#"&filha
					end if
					
					conta_filhas=conta_filhas+1	
				RSNNa.MOVENEXT
				wend	
			end if	
			
			'media_aluno=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, nu_matricula, disciplina, vetor_filhas, CAMINHOn, notaFIL, periodo)	
			media_aluno=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, nu_matricula, disciplina, vetor_filhas, CONn, notaFIL, periodo, "VA_Media2", outro)				
			if media_aluno="&nbsp;" then
				pula_aluno="S"				
			else
				media_denominador=media_denominador+1
			end if		
		else
			media_aluno=0
			media_denominador=media_denominador	
		end if	
		
		if media_aluno="&nbsp;" or media_aluno="" or isnull(media_aluno) then
			media_aluno=0	
			pula_aluno="S"						
		end if	
		media_numerador=media_numerador*1
		media_aluno=media_aluno*1
		acumula_aluno=acumula_aluno+media_aluno
	NEXT
		if pula_aluno="N" then
			media_numerador=acumula_aluno			

			if media_denominador=0 then
				media=media_numerador
			else
				media=media_numerador/media_denominador		
			end if
	
			media=formatnumber(media,1)
			
			if nu_media_check = 1 then
				vetor_medias=media
				vetor_aluno_media=nu_matricula
			else
				vetor_medias=vetor_medias&";"&media
				vetor_aluno_media=vetor_aluno_media&";"&nu_matricula
			end if
			nu_media_check=nu_media_check+1			
		end if				
	RSt0.MOVENEXT	
wend




faixa1=0
faixa2=0
faixa3=0
faixa4=0
faixa5=0

vetor_medias=split(vetor_medias,";")
vetor_aluno_media=split(vetor_aluno_media,";")

for n=0 to ubound(vetor_medias)
	analisa_media=vetor_medias(n)
	analisa_media=analisa_media*1

	if analisa_media>80 then
		faixa5=faixa5+1
		alunos_faixa5=alunos_faixa5&"#!#"&vetor_aluno_media(n)
		medias_faixa5=medias_faixa5&"#!#"&analisa_media
		
	elseif analisa_media>60 then
		faixa4=faixa4+1
		alunos_faixa4=alunos_faixa4&"#!#"&vetor_aluno_media(n)
		medias_faixa4=medias_faixa4&"#!#"&analisa_media

	elseif analisa_media>40 then
		faixa3=faixa3+1
		alunos_faixa3=alunos_faixa3&"#!#"&vetor_aluno_media(n)
		medias_faixa3=medias_faixa3&"#!#"&analisa_media

	elseif analisa_media>20 then
		faixa2=faixa2+1
		alunos_faixa2=alunos_faixa2&"#!#"&vetor_aluno_media(n)
		medias_faixa2=medias_faixa2&"#!#"&analisa_media

	else
		faixa1=faixa1+1
		alunos_faixa1=alunos_faixa1&"#!#"&vetor_aluno_media(n)
		medias_faixa1=medias_faixa1&"#!#"&analisa_media

	end if					

next
session("faixas")=faixa1&"#!#"&faixa2&"#!#"&faixa3&"#!#"&faixa4&"#!#"&faixa5
session("categorias")="0-20#!#21-40#!#41-60#!#61-80#!#81-100"
%>

<table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>    
  <th scope="col" class="tb_tit">Notas</th>
<%

faixas=session("faixas")
categorias=session("categorias")

classes=split(categorias,"#!#")
'response.Write(ubound(classes))

for y=0 to ubound(classes)
	'response.Write(classes(y)&"<BR>")
	nomes = classes(y)
	faixa=y
	faixa=faixa*1
	faixa_origem=faixa_origem*1
	
	
	if faixa_origem=0 then
		alunos_pesquisados=alunos_faixa1
		medias_pesquisadas=medias_faixa1		
		
	elseif faixa_origem=1 then	
		alunos_pesquisados=alunos_faixa2
		medias_pesquisadas=medias_faixa2	

	elseif faixa_origem=2 then	
		alunos_pesquisados=alunos_faixa3	
		medias_pesquisadas=medias_faixa3
			
	elseif faixa_origem=3 then	
		alunos_pesquisados=alunos_faixa4
		medias_pesquisadas=medias_faixa4
		
	else	
		alunos_pesquisados=alunos_faixa5
		medias_pesquisadas=medias_faixa5
		
	end if			
	if faixa=faixa_origem then
%>
        <th scope="col" class="tb_subtit"><%response.Write(nomes)%></th>
	<%else%>  
        <th scope="col" class="tb_subtit"><a href="detalhar.asp?opt=grafico&fx=<%response.Write(faixa)%>&obr=<%response.Write(obr)%>&order=<%response.Write(order)%>"><%response.Write(nomes)%></a></th>  
<%
	end if
next%>
  </tr>
  <tr>
      <td align="center" class="tb_tit">Qtde/Alunos</td>
 <%
 valores=split(faixas,"#!#")
for i=0 to ubound(valores)
Vals= valores(i)
%>
    <td align="center" class="form_corpo"><%response.Write(Vals)%></td>
<%
next
%>  
  </tr>
</table>
</td>
                  </tr>
                  <tr>
                    <td align="center" valign="top">&nbsp;</td>
                  </tr>
                  <tr>
                    <td align="center" valign="top">
<% aluno_ordena=split(alunos_pesquisados,"#!#")
 medias_ordena=split(medias_pesquisadas,"#!#")
if ubound(aluno_ordena)=-1 then
 %>
 <div align="center"><font class="form_corpo"> Não existem alunos nessa faixa de notas</font></div>
<%else
' ordena vetor médias
	AuxVetor = medias_ordena
	AuxVetor_aluno = aluno_ordena
	for x=1 to ubound(medias_ordena)
		Valor=0 
		menor_media=101
		maior_media=0
		if order="d" then
			for y=1 to ubound(AuxVetor)  
				media_teste=AuxVetor(y)
				media_teste=media_teste*1
				maior_media=maior_media*1
				Valor=Valor*1
				y=y*1
	
					if media_teste>=maior_media and y<>Valor then
						maior_media=media_teste
						aluno_maior_media=AuxVetor_aluno(y)
						Valor=y
					end if			
			next
			media_ordenada=media_ordenada&"#!#"&maior_media
			aluno_ordenada=aluno_ordenada&"#!#"&aluno_maior_media
			
		elseif order="c" then	
			for y=1 to ubound(AuxVetor)  
				media_teste=AuxVetor(y)
				media_teste=media_teste*1
				menor_media=menor_media*1
				maior_media=maior_media*1
				Valor=Valor*1
				y=y*1
	
					if media_teste<=menor_media and y<>Valor then
						menor_media=media_teste
						aluno_menor_media=AuxVetor_aluno(y)
						Valor=y
					end if
			next
			media_ordenada=media_ordenada&"#!#"&menor_media
			aluno_ordenada=aluno_ordenada&"#!#"&aluno_menor_media
		end if
		

' Retirando o menor ou maior elemento do vetor das médias
			Dim tmpvetor
			tmpvetor = array()
			for z = LBound( AuxVetor ) to UBound ( AuxVetor )
				z=z*1
				Valor=Valor*1
				if z <> Valor then
					Redim preserve tmpvetor ( UBound(tmpvetor)+1 ) 
					tmpvetor ( UBound ( tmpvetor ) ) = AuxVetor( z )
				end if
			next
			AuxVetor = tmpvetor 'salvando novamente a Array
			tmpvetor = array() 'liberando a var tmp 

' Retirando o dono da menor ou maior média do vetor dos alunos
			Dim tmpvetor_aluno
			tmpvetor_aluno = array()

			for z = LBound( AuxVetor_aluno) to UBound ( AuxVetor_aluno )
				z=z*1
				Valor=Valor*1
				if z <> Valor then
					Redim preserve tmpvetor_aluno ( UBound(tmpvetor_aluno)+1 ) ' adicionei um elemento
					tmpvetor_aluno ( UBound ( tmpvetor_aluno ) ) = AuxVetor_aluno( z )
				end if
			next
			AuxVetor_aluno = tmpvetor_aluno 'salvando novamente a Array
			tmpvetor_aluno = array() 'liberando a var tmp 


	next
medias_exibe=split(media_ordenada,"#!#")
aluno_exibe=split(aluno_ordenada,"#!#")

%>                     
        <table width="800" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <th scope="col" width="100" class="tb_subtit">Turma</th> 
            <th scope="col" width="30" class="tb_subtit">Nº</th>   
            <th width="555" align="left" class="tb_subtit" scope="col">Nome</th>
            <th width="15" align="left" class="tb_subtit" scope="col">
           <%if order="d"then%>
           <img src="../../../../img/decres01.gif" width="12" height="25">
           <%else%>
            <a href="detalhar.asp?opt=grafico&fx=<%response.Write(faixa_origem)%>&obr=<%response.Write(obr)%>&order=d"><img src="../../../../img/decres02.gif" width="12" height="25" border="0"></a>
            <%end if%>
            </th>
            <th width="15" align="left" class="tb_subtit" scope="col">
            <%if order="d"then%>
            <a href="detalhar.asp?opt=grafico&fx=<%response.Write(faixa_origem)%>&obr=<%response.Write(obr)%>&order=c"><img src="../../../../img/cres02.gif" width="12" height="25" border="0"></a>           
			<%else%>
           <img src="../../../../img/cres01.gif" width="12" height="25"> 
            <%end if%>
            </th>
            <th scope="col" width="80" class="tb_subtit">Média-Geral</th>
          </tr>
          <%
      
        check = 1 
        for k=1 to ubound(aluno_exibe)  
        
                Set RSnome = Server.CreateObject("ADODB.Recordset")
                SQLnome = "SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.NU_Ano ="& ano_letivo &" AND TB_Matriculas.CO_Matricula ="& aluno_exibe(k) 
                RSnome.Open SQLnome, CON4
                
        turma= RSnome("CO_Turma")
        nu_chamada = RSnome("NU_Chamada")
        nome_aluno = RSnome("NO_Aluno")
        media_exibe=formatnumber(medias_exibe(k),1)
        
            if check mod 2 =0 then
                cor = "tb_fundo_linha_par" 
            else
                cor ="tb_fundo_linha_impar"
            end if
          %> 
            <tr class="<%response.Write(cor)%>">
            <td><div align="center"><%response.Write(turma)%></div></td>    
            <td><div align="center"><%response.Write(nu_chamada)%></div></td>
            <td colspan="3"><%response.Write(nome_aluno)%></td>
            <td><div align="center"><%response.Write(media_exibe)%></div></td>
            </tr>
        <%
        check=check+1
        Next
        %> 
     </table>
<%end if%>     
     </td></tr>
                  <tr>
                    <td align="center" valign="top">&nbsp;</td>
                  </tr>     
                  <tr>
                    <td align="center" valign="top"><hr></td>
                  </tr>
                  <tr>
                    <td align="center" valign="top"><table width="1000" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="33%" align="center"><input type="button" name="Voltar" id="Voltar" value="Voltar" class="botao_cancelar" onClick="MM_goToURL('parent','mapa.asp?opt=vt&obr=<%response.Write(obr)%>');return document.MM_returnValue" ></td>
                        <td width="34%" align="center">&nbsp;</td>
                        <td width="33%" align="center">&nbsp;</td>
                      </tr>
                    </table></td>
                  </tr>
        </table>
              </form></td>
  </tr>
  <tr>
    <td height="40"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

</body>
<%
call GravaLog (chave,obr)
%>
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
pasta=arPath(seleciona)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>