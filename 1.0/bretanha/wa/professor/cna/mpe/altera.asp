<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->


<!--#include file="../../../../inc/caminhos.asp"-->



<%opt = request.QueryString("opt")
nvg = session("chave")
co_usr = session("co_user")
chave=nvg
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
nivel=4
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

total_datas = 5

if opt = "err1" OR opt = "err2" OR opt = "err3"then
dd=request.querystring("dd")
dados = split(dd,"_")
ano_letivo = dados(0)
curso = dados(1)
unidade = dados(2)
co_grupo = dados(3)
else

co_grupo = request.Form("co_grupo")
curso = request.Form("curso")
unidade = request.Form("unidade")
co_etapa= request.form("etapa")
turma= request.form("turma")
end if

if isnull(unidade) or unidade="" then
unidade=Session("unidade")
else
Session("unidade")=unidade
end if

if isnull(curso) or curso="" then
curso=Session("curso")
else
Session("curso")=curso
end if

if isnull(co_etapa) or co_etapa="" then
co_etapa=Session("co_etapa")
else
Session("co_etapa")=co_etapa
end if

if isnull(turma) or turma="" then
turma=Session("turma")
else
Session("turma")=turma
end if

nivel=4



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

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Abreviado_Curso")

 call navegacao (CON,chave,nivel)
navega=Session("caminho")

grava_nota= unidade&"?"&curso&"?"&co_etapa&"?"&turma
session("grava_nota")=grava_nota

onLoad="onLoad=""redimensiona()"""	
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
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresiz!=MM_reloadPage; }}
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
                                                            alert("Esse browser n�o tem recursos para uso do Ajax");
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
document.all.divTurma.innerHTML = "<select class=select_style></select>"
//recuperarEtapa()
                                                           }
                                               }

                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=e", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select class=select_style></select>"
//recuperarTurma()
                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t2", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t
																	   
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
<body link="#6699CC" vlink="#6699CC" alink="#6699CC" leftmargin="0" background="../../../../img/fundo.gif" topmargin="0" marginwidth="0" marginheight="0" <%response.Write(onload)%>>
<% call cabecalho (nivel) %>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF" bgimage="../../../../fundo_interno.gif">
  <tr>     
    <td height="10" valign="top" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)
%></font> </td>
  </tr>
      <%if opt = "err1" then%>
  <tr> 
    <td height="10" valign="top"> 
      <%
		call mensagens(nivel,231,1,0)
%>
    </td>
                </tr>
                <%end if%>
                <%if opt = "err2" then%>
                <tr> 
                  
    <td height="10" valign="top"> 
      <%
		call mensagens(nivel,232,1,0)
%>
    </td>
                </tr>
                <%end if%>
                <%if opt = "err3" then%>
                <tr> 
                  
    <td height="10" valign="top"> 
      <%
		call mensagens(nivel,233,1,0)
%>
    </td>
                </tr>
                <%end if%>
<tr height="10">                  
    <td height="10">
    <DIV ID="MSG1"><%	call mensagens(nivel,11,0,0) %></DIV>
	<DIV ID="MSG2" style="display:none"><%	call mensagens(nivel,670,0,total_datas) %></DIV>    
    </td>
                </tr>
            <td valign="top"> 
<form name="inclusao" method="post" action="altera.asp">
                
        <table width="1000" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td height="15" class="tb_tit">Grade de Aulas
<input name="co_grupo" type="hidden" id="co_grupo" value="<% = co_grupo %>"></td>
          </tr>
          <tr> 
            <td valign="top">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="250" class="tb_subtit"> <div align="center">UNIDADE 
                    </div></td>
                  <td width="250" class="tb_subtit"> <div align="center">CURSO 
                    </div></td>
                  <td width="250" class="tb_subtit"> <div align="center">ETAPA 
                    </div></td>
                  <td width="250" class="tb_subtit"> <div align="center">TURMA 
                    </div></td>
                </tr>
                <tr> 
                  <td> <div align="center"> 
                      <select name="unidade" class="select_style" onchange="recuperarCurso(this.value)">
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
                  <td> <div align="center"> 
                      <div id="divCurso"> 
                        <select name="curso" class="select_style" onchange="recuperarEtapa(this.value)">
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
                  <td> <div align="center"> 
                      <div id="divEtapa"> 
                        <select name="etapa" class="select_style" onchange="recuperarTurma(this.value)">
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
                  <td> <div align="center"> 
                      <div id="divTurma"> 
                        <select name="turma" class="select_style" onChange="MM_callJS('submitfuncao()')">
                          <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Turma where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"' AND CO_Etapa='" & co_etapa & "' order by CO_Turma" 
		RS3.Open SQL3, CON0						

while not RS3.EOF
co_turma= RS3("CO_Turma")

if co_turma=turma then
%>
                          <option value="<%response.Write(co_turma)%>" selected> 
                          <%response.Write(co_turma)%>
                          </option>
                          <%
else
%>
                          <option value="<%=co_turma%>"> 
                          <%response.Write(co_turma)%>
                          </option>
                          <%
co_turma_check = co_turma
end if
RS3.MOVENEXT
WEND
%>
                        </select>
                      </div>
                    </div></td>
                </tr>
                <tr> 
                  <td colspan="4">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="4"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="350" class="tb_subtit"> <div align="center">DISCIPLINA</div></td>
                        <td width="400" class="tb_subtit"> 
                          <div align="center">PROFESSOR</div></td>
                        <td width="50" class="tb_subtit"> <div align="center">B1</div></td>
                        <td width="50" class="tb_subtit"> <div align="center">B2</div></td>
                        <td width="50" class="tb_subtit"> <div align="center">B3</div></td>
                        <td width="50" class="tb_subtit"> <div align="center">B4</div></td>
                        <td width="50" class="tb_subtit"> <div align="center">REC</div></td>
                      </tr>
                      
                      <%
conta_grid = 0
vetor_grid = ""			

'		Set RS5 = Server.CreateObject("ADODB.Recordset")
'		SQL5 = "SELECT TB_Programa_Aula.CO_Curso, TB_Programa_Aula.CO_Etapa, TB_Programa_Aula.CO_Materia, TB_Programa_Subs.CO_Materia_Filha, TB_Programa_Aula.NU_Ordem_Boletim, TB_Programa_Subs.NU_Ordem_Boletim FROM (TB_Materia INNER JOIN TB_Programa_Aula ON TB_Materia.CO_Materia = TB_Programa_Aula.CO_Materia) LEFT JOIN TB_Programa_Subs ON (TB_Programa_Aula.CO_Materia = TB_Programa_Subs.CO_Materia_Principal) AND (TB_Programa_Aula.CO_Etapa = TB_Programa_Subs.CO_Etapa) AND (TB_Programa_Aula.CO_Curso = TB_Programa_Subs.CO_Curso) where TB_Programa_Aula.CO_Etapa ='"& co_etapa &"' AND TB_Programa_Aula.CO_Curso ='"& curso &"'  ORDER BY TB_Programa_Aula.CO_Curso, TB_Programa_Aula.CO_Etapa, TB_Programa_Aula.NU_Ordem_Boletim, TB_Programa_Subs.NU_Ordem_Boletim"
'		RS5.Open SQL5, CON0	

		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim"
		RS5.Open SQL5, CON0

while not RS5.EOF
'co_mat_sub= RS5("CO_Materia_Filha")
co_mat_prin = RS5("CO_Materia")
'if isnull(co_mat_sub) then
'	co_mat_sub= RS5("CO_Materia")
'end if	

'response.Write(co_mat_prin&"====================<br>")

	Set RS5a = Server.CreateObject("ADODB.Recordset")
	SQL5a = "SELECT * FROM TB_Programa_Subs where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia_Principal ='"& co_mat_prin &"' order by NU_Ordem_Boletim "
	RS5a.Open SQL5a, CON0
				
	if not RS5a.EOF then
		conta_linha_sub = 0		
		while not RS5a.EOF
			co_mat_sub= RS5a("CO_Materia_Filha")	
			'response.Write(conta_linha_sub&" "&co_mat_sub&"<BR>")
			vetor_grid_linha_sub = gera_grid_linha(co_mat_prin, co_mat_sub, unidade, curso, co_etapa, turma)	
			
			if conta_linha_sub=0 then
				vetor_grid_linha = vetor_grid_linha_sub		
			else
				
				vetor_grid_linha = vetor_grid_linha&"$!$"&vetor_grid_linha_sub	
			end if	
			
			conta_linha_sub = conta_linha_sub+1				
					
		RS5a.MOVENEXT
		WEND	
	else	
	
		vetor_grid_linha = gera_grid_linha(co_mat_prin, "", unidade, curso, co_etapa, turma)		
	
	end if
'response.Write(vetor_grid_linha&"<BR>")	

	if conta_grid=0 then
		vetor_grid = vetor_grid_linha		
	else
		
		vetor_grid = vetor_grid&"$!$"&vetor_grid_linha	
	end if	
	
	conta_grid = conta_grid+1	
	
RS5.MOVENEXT
WEND	
'response.Write("=========================<BR>"&vetor_grid&"<BR>")	
'RESPONSE.End()
check=0	
split_grid = split(vetor_grid,"$!$")
for s=0 to ubound(split_grid)

	 if check mod 2 =0 then
	  cor = "tb_fundo_linha_par" 
	 else cor ="tb_fundo_linha_impar"
	  end if

%>
    <tr>
    <%
	split_linha = split(split_grid(s),"#!#")	
	for l=0 to ubound(split_linha)
		largura = 50
		if l=0 then
			largura = 300		
		elseif l=1 then
			largura = 450			
		end if 	
	%>
	<td width="<%response.write(largura)%>" class="<%=cor%>" align="center"> 
	  <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
	  		<%
				texto = split_linha(l)			
				split_professor = split(split_linha(l),"#$#")
				if l=0 then
					co_mat=split_professor(0)
					texto=split_professor(1)
				end if
				if l=1 then
					cod_professor = split_professor(0)
					if split_professor(2) = "FALSO" then
						classe = "inativo"
					else
						classe = "ativo"						
					end if
					texto = "<a class="&classe&" href=../../man/professores/altera.asp?ori=01&opt=ln&cod_cons="&split_professor(0)&"&nvg=WA-PF-MA-APR target=_parent>"&split_professor(1)&"</a>"
				end if
				if l>1 then
					if texto="x" then
						imagem = "s.gif"
					else
						imagem = "n.gif"				
					end if	
					texto = "<a href='notas.asp?d="&co_mat&"&pr="&cod_professor&"&exb="&total_datas&"&p="&l-1&"'><img src='../../../../img/"&imagem&"' width='8' height='8' border='0'></a>"
				end if
				response.Write(texto)
				
				
			 %>
      </font>
     </td>
	<%
	next
    %>
    </tr>
<%	
check=check+1	
next		
%>
              </table>
            </td>
          </tr>
        </table>
              </td>
          </tr>
</table>
</form>
  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
</body>
<%

'end if
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
pasta=arPath(seleciona1)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>