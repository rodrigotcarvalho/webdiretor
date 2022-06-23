<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/boletim.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%
opt=request.QueryString("obr")
ori=request.QueryString("ori")
co_prof = request.QueryString("cod_cons")
co_materia=request.QueryString("d")
autoriza=Session("autoriza")
Session("autoriza")=autoriza

permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("chave")
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
nivel=4


'if ori="1" then
co_materia = request.form("mat_prin")
unidade= request.form("unidade")
curso= request.form("curso")
co_etapa= request.form("etapa")
turma= request.form("turma")
ano_letivo = session("ano_letivo")
co_usr = session("co_usr")
co_prof = request.form("co_prof")



'else

'dados= split(opt, "-" )
'unidade= dados(0)
'curso= dados(1)
'co_etapa= dados(2)
'turma= dados(3)
'co_prof = request.QueryString("cod_cons")
'co_materia=request.QueryString("d")
'co_usr = session("co_usr")
'end if

'response.Write(co_materia&"<<")
obr=co_materia&"_"&unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&ano_letivo&"_"&co_prof
'response.Write(obr)


if etapa = "999999" then
		response.redirect("tabelas2.asp?opt=err1&or=1&dd="&ano_letivo&"_"&curso&"_"&unidade&"_"&co_prof)
ELSEif co_materia = "999999" then
		response.redirect("tabelas2.asp?opt=err2&or=1&dd="&ano_letivo&"_"&curso&"_"&unidade&"_"&co_prof)
ELSEif turma = "999999" then
	response.redirect("tabelas2.asp?opt=err3&or=1&dd="&ano_letivo&"_"&curso&"_"&unidade&"_"&co_prof)

else

'obr=unidade&"_"&curso&"_"&etapa&"_"&turma&"_"&ano_letivo&"_"&co_materia

		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

 call navegacao (CON,chave,nivel)
navega=Session("caminho")

		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"'"
		Set RS = CON2.Execute(CONEXAO)
if RS.EOF then
response.Write("<div align=center><font size=2 face=Courier New, Courier, mono  color=#990000><b>Esta turma não está disponível no momento</b></font><br")
response.Write("<font size=2 face=Courier New, Courier, mono  color=#990000><a href=javascript:window.history.go(-1)>voltar</a></font></div>")

else
tb_nota = RS("TP_Nota")
end if


%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../../../estilos.css" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_popupMsg(msg) { //v1.0
  alert(msg);
}
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

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
function submitfuncao()  
{
   var f=document.forms[0]; 
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

                                               oHTTPRequest.open("post", "../../../../inc/executa_wn.asp?opt=c", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
document.all.divEtapa.innerHTML ="<select class=select_style></select>"
document.all.divTurma.innerHTML = "<select class=select_style></select>"
document.all.divDisciplima.innerHTML = "<select class=select_style></select>"
//recuperarEtapa()
                                                           }
                                               }

                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa_wn.asp?opt=e", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select class=select_style></select>"
document.all.divDisciplima.innerHTML = "<select class=select_style></select>"
//recuperarTurma()
                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa_wn.asp?opt=t3", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t
document.all.divDisciplima.innerHTML = "<select class=select_style></select>"																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
						 function recuperarDisciplina(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa_wn.asp?opt=d3", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_d= oHTTPRequest.responseText;
resultado_d = resultado_d.replace(/\+/g," ")
resultado_d = unescape(resultado_d)
document.all.divDisciplima.innerHTML = resultado_d																	   
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

<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0">
<form name="inclusao" method="post" action="mapa.asp?ori=1"> 
<% call cabecalho(nivel) %>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr>                    
            <td height="10" class="tb_caminho"> <font class="style-caminho">
              <%
	  response.Write(navega)

%>
              </font>
	</td>
  </tr><%
	no_materia= GeraNomes("D",co_materia,variavel2,variavel3,variavel4,variavel5,CON0,outro) 	
	no_unidade= GeraNomes("U",unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)
	no_curso=GeraNomes("C",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro) 	
	no_etapa=GeraNomes("E",curso,co_etapa,variavel3,variavel4,variavel5,CON0,outro)


nome_prof = session("nome_prof") 
tp=	session("tp")

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("m", now) 
data = dia &"/"& mes &"/"& ano
horario = hora & ":"& min
acesso_prof = session("acesso_prof")

if co_prof="&nbsp;" then
		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"' AND CO_Materia = '"& co_materia &"'"
		Set RS = CON2.Execute(CONEXAO)
else
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE CO_Professor= "& co_prof &"AND NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"' AND CO_Materia = '"& co_materia &"'"
		Set RS = CON2.Execute(CONEXAO)
end if

%><tr>              
    <td height="10"> 
      <%	call mensagens(nivel,636,0,0) %>
    </td>
                </tr>
<td height="15" class="tb_tit">Grade de Aulas 

          </td>
                  </tr>
                  <tr> 
                    
    <td height="10">
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="200" class="tb_subtit"> <div align="center">UNIDADE 

            </div></td>
            <td width="200" class="tb_subtit"> 
              <div align="center">CURSO </div></td>
            <td width="200" class="tb_subtit"> 
              <div align="center">ETAPA</div></td>
            <td width="200" class="tb_subtit"> 
              <div align="center">TURMA </div></td>
            <td width="200" class="tb_subtit"> 
              <div align="center">DISCIPLINA</div></td>
        </tr>
        <tr>
            <td width="200"> 
              <div align="center">               <input name="co_prof" type="hidden" id="co_prof" value="<%response.write(co_prof)%>">
                      <select name="unidade" class="select_style" onChange="recuperarCurso(this.value)">
                        <option value="0"></option>
                        <%		
		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0
NU_Unidade_Check=999999		
While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")

		Set RSUn_Prof = Server.CreateObject("ADODB.Recordset")
		SQLUn_Prof = "SELECT Distinct NU_Unidade FROM TB_Da_Aula Where CO_Professor="&co_prof&" AND NU_Unidade="& NU_Unidade &" ORDER BY NU_Unidade"
		RSUn_Prof.Open SQLUn_Prof, CON2

if RSUn_Prof.eof then		
RS0.MOVENEXT
else	
NU_Unidade=NU_Unidade*1
unidade=unidade*1
if NU_Unidade = unidade then
%>
                        <option value="<%response.Write(NU_Unidade)%>" selected>
              <%response.Write(NO_Abr)%></option>
              <%else
%>
              <option value="<%response.Write(NU_Unidade)%>"> 
              <%response.Write(NO_Abr)%>
              </option>
              <%

end if
RS0.MOVENEXT
end if
WEND
%></select>
              </div></td>
          <td width="200"> <div align="center"> 
              <div id="divCurso"> 
                <select name="curso" class="select_style" onChange="recuperarEtapa(this.value)">
                  <%		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT DISTINCT CO_Curso FROM TB_Da_Aula Where CO_Professor="&co_prof&" AND NU_Unidade="& unidade
		RS0.Open SQL0, CON2
		
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
          <td width="200"> <div align="center"> 
              <div id="divEtapa"> 
                <select name="etapa" class="select_style" onChange="recuperarTurma(this.value)">
                  <%		

		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Da_Aula Where CO_Professor="&co_prof&" AND NU_Unidade="& unidade &" AND CO_Curso='"&curso&"'"
		RS0b.Open SQL0b, CON2
		
		
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
          <td width="200"> <div align="center"> 
              <div id="divTurma"> 
                <select name="turma" class="select_style" onChange="recuperarDisciplina()">
                  <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Da_Aula where CO_Professor="&co_prof&" AND NU_Unidade="&unidade&"AND CO_Curso='"&curso&"' AND CO_Etapa='" & co_etapa & "' order by CO_Turma" 
		RS3.Open SQL3, CON2						

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
          <td width="200"> <div id="divDisciplima" align="center"> 
              <select name="mat_prin" class="select_style" onChange="MM_callJS('submitfuncao()')">
                <%
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT DISTINCT CO_Materia_Principal FROM TB_Da_Aula where CO_Professor ="& co_prof&" and CO_Etapa = '"& co_etapa &"' AND NU_Unidade = "&unidade&" and CO_Curso = '"& curso &"' order by CO_Materia_Principal"
		RS5.Open SQL5, CON2

while not RS5.EOF
co_mat_prin= RS5("CO_Materia_Principal")


		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Materia where CO_Materia ='"& co_mat_prin &"'"
		RS7.Open SQL7, CON0
		
		no_mat_prin= RS7("NO_Materia")
		
if co_materia=co_mat_prin then
%>
                <option value="<%=co_mat_prin%>" selected> 
                <%response.Write(no_mat_prin)%>
                </option>
                <%
else
%>
                <option value="<%=co_mat_prin%>"> 
                <%response.Write(no_mat_prin)%>
                </option>
                <%
end if
RS5.MOVENEXT
WEND%>
                
              </select>
            </div></td>
        </tr>
      </table>
	  </td>
                  </tr>
                  <tr bgcolor="#FFFFFF"> 
                    
    <td height="10" bgcolor="#FFFFFF">&nbsp;</td>
                <tr> 
            <td valign="top"> 
              <div align="right"> 
                <%

	if tb_nota ="TB_NOTA_A" then
	CAMINHOn = CAMINHO_na
	
	elseif tb_nota="TB_NOTA_B" then
		CAMINHOn = CAMINHO_nb
	
	elseif tb_nota ="TB_NOTA_C" then
			CAMINHOn = CAMINHO_nc
			
	elseif tb_nota ="TB_NOTA_D" then
			CAMINHOn = CAMINHO_nd

	elseif tb_nota ="TB_NOTA_E" then
			CAMINHOn = CAMINHO_ne
						
	else
			response.Write("ERRO")
	end if
	
	call boletim_escolar (unidade,curso,co_etapa,turma,CAMINHOn,tb_nota,co_materia,"MRD")	
 
    Set RS = Nothing
	    Set RS2 = Nothing
		    Set RS3 = Nothing
End if

call GravaLog (chave,obr)
 %>
                <hr>
              </div></td>
  </tr>
  <tr>
    <td height="40"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
</form>
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
pasta=arPath(seleciona)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>