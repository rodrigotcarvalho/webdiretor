<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes7.asp"-->
<%
opt = REQUEST.QueryString("opt")
obr = request.QueryString("obr")
nivel=4

autoriza=Session("autoriza")
Session("autoriza")=autoriza

permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("chave")
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

if opt= "vt" then
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
end if

ano_letivo = session("ano_letivo")
obr=unidade&"_"&curso&"_"&co_etapa&"_"&periodo&"_"&ano_letivo


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
function submitfuncao()  
{
   var f=document.forms[0]; 
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
<body background="../../../../img/fundo.gif"leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="inclusao" method="post" action="mapa.asp" onSubmit="return checksubmit()">
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
                <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
                  <tr class="tb_tit"
> 
                    <td width="653" height="15" class="tb_tit"
>Segmento </td>
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
		tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")				
			
			
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo where TP_Modelo='"&tp_modelo&"' order by NU_Periodo"
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
else
		response.Write("ERRO")
end if
		Set CONn = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONn.Open ABRIRn


		Set RSNN = Server.CreateObject("ADODB.Recordset")
		CONEXAONN = "Select * from TB_Programa_Aula WHERE CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa&"' order by NU_Ordem_Boletim"
		Set RSNN = CON0.Execute(CONEXAONN)	
				
nu_materia_check = 1
while not RSNN.EOF
co_materia= RSNN("CO_Materia")

	if nu_materia_check = 1 then
		vetor_materias="'"&co_materia&"'"
	else
		vetor_materias=vetor_materias&",'"&co_materia&"'"
	end if
nu_materia_check=nu_materia_check+1	
RSNN.MOVENEXT
wend

		Set RSt0 = Server.CreateObject("ADODB.Recordset")
		SQLt0 = "SELECT * FROM TB_Matriculas where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND NU_Ano ="&ano_letivo&" And CO_Situacao='C' order by NU_Chamada"
		RSt0.Open SQLt0, CON4
		
nu_media_check = 1
while not RSt0.EOF
nu_matricula = RSt0("CO_Matricula")

		Set RSnFIL = Server.CreateObject("ADODB.Recordset")
		SQLnFIL = "SELECT SUM("&notaFIL&"."&m_cons&")AS sumVA_Media3 FROM "&notaFIL&" where CO_Matricula ="&nu_matricula&" and CO_Materia in ("& vetor_materias&") and NU_Periodo="&periodo
		RSnFIL.Open SQLnFIL, CONn

	IF RSnFIL.EOF then
		RSt0.MOVENEXT
	else
		soma_aluno=RSnFIL("sumVA_Media3")
		if isnull(soma_aluno) or soma_aluno="" then
		else
		media_aluno=soma_aluno/nu_materia_check
			media_aluno=formatnumber(media_aluno,1)
'			response.Write("media_aluno="&media_aluno&"<BR>")
				if nu_media_check = 1 then
					vetor_medias=media_aluno
				else
					vetor_medias=vetor_medias&";"&media_aluno
				end if
			nu_media_check=nu_media_check+1		
'			response.Write("vetor_medias="&vetor_medias&"<BR>")
		end if	
		RSt0.MOVENEXT
	end if
wend

faixa1=0
faixa2=0
faixa3=0
faixa4=0
faixa5=0

vetor_medias=split(vetor_medias,";")

for n=0 to ubound(vetor_medias)
	analisa_media=vetor_medias(n)
'response.Write(analisa_media&"<<BR>")
	analisa_media=analisa_media*1

	if analisa_media>80 then
		faixa5=faixa5+1
	elseif analisa_media>60 then
		faixa4=faixa4+1
	elseif analisa_media>40 then
		faixa3=faixa3+1
	elseif analisa_media>20 then
		faixa2=faixa2+1
	else
		faixa1=faixa1+1
	end if					
next
'response.end()
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

	%>
    <th scope="col" class="tb_subtit"><a href="detalhar.asp?opt=grafico&fx=<%response.Write(faixa)%>&obr=<%response.Write(obr)%>&order=d"><%response.Write(nomes)%></a></th>
<%next%>
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
<DIV align="center">
<iframe src ="iframe.asp" frameborder ="0" width="1000" height="400" align="middle"> </iframe>
</DIV>
</td>
                  </tr>
                </table>
              </td>
  </tr>
  <tr>
    <td height="40"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
</form>
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