<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes4.asp"-->
<!--#include file="../../../../../global/conta_alunos.asp"-->

<%
opt=request.QueryString("opt")
obr = request.QueryString("obr")
autoriza=session("autoriza")
nvg = session("chave")
ano_letivo = session("ano_letivo")
co_usr = session("co_user")
grupo=session("grupo")
chave=nvg
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
nivel=4
trava=session("trava")
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)



		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		

if opt = "err6" then

hp= request.QueryString("hp")

dados= split(obr, "?" )
unidade= dados(0)
curso= dados(1)
co_etapa= dados(2)
turma= dados(3)
ano_letivo = dados(4)
co_usr = session("co_usr")


alt= split(hp, "_" )
calc=alt(1)
qerrou= alt(2)
errou = alt(3)

if calc=0 then

else
				Set CONer= Server.CreateObject("ADODB.Connection") 
		ABRIRer = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONer.Open ABRIRer
		
		Set RSer  = Server.CreateObject("ADODB.Recordset")
		SQL_er  = "Select * from TB_Aluno_Esta_Turma WHERE CO_Matricula = "& calc
		Set RSer  = CONer.Execute(SQL_er )
		


errante = RSer("NU_Chamada")
end if

select case errou
case "f1"
errado ="B1"
case "f2"
errado ="B2"
case "f3"
errado ="B3"
case "f4"
errado ="B4"
end select

elseif opt="ok" or  opt= "vt" or opt="cln" then
dados= split(obr, "?" )
unidade= dados(0)
curso= dados(1)
co_etapa= dados(2)
turma= dados(3)
ano_letivo = dados(4)
co_usr = session("co_usr")

elseif opt="cgp" then
unidade= request.QueryString("u")
curso= request.QueryString("c")
co_etapa= request.QueryString("e")
turma= request.QueryString("t")
co_usr = session("co_usr")

else
grava_nota=session("grava_nota")

co_grupo = request.Form("co_grupo")
curso = request.Form("curso")
unidade = request.Form("unidade")
co_etapa= request.form("etapa")
turma= request.form("turma")

co_usr = session("co_usr")
end if

session("unidade")=unidade
session("curso")=curso
session("co_etapa")=co_etapa
session("turma")=turma



'response.Write(unidade&" - "&curso&" - "&co_etapa&" - "&turma)

if co_etapa = "999999" then
		response.redirect("tabelas2.asp?opt=err1&or=1&dd="&ano_letivo&"_"&curso&"_"&unidade&"_"&co_prof)
ELSEif co_materia = "999999" then
		response.redirect("tabelas2.asp?opt=err2&or=1&dd="&ano_letivo&"_"&curso&"_"&unidade&"_"&co_prof)
ELSEif turma = "999999" then
	response.redirect("tabelas2.asp?opt=err3&or=1&dd="&ano_letivo&"_"&curso&"_"&unidade&"_"&co_prof)

else

obr=unidade&"?"&curso&"?"&co_etapa&"?"&turma&"?"&ano_letivo

				Set CONg = Server.CreateObject("ADODB.Connection") 
		ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONg.Open ABRIRg

				Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"'"
		Set RS = CONg.Execute(CONEXAO)

'response.Write(CONEXAO)
if RS.EOF then
response.Write("<div align=center><font size=2 face=Courier New, Courier, mono  color=#990000><b>Esta turma não está disponível no momento</b></font><br")
response.Write("<font size=2 face=Courier New, Courier, mono  color=#990000><a href=javascript:window.history.go(-1)>voltar</a></font></div>")

else
nota = RS("TP_Nota")
coordenador = RS("CO_Cord")
end if

 call navegacao (CON,chave,nivel)
navega=Session("caminho")
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
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
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

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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

function mudar_cor_focus(celula){
   celula.style.backgroundColor="#dddddd"

}
function mudar_cor_blur_par(celula){
   celula.style.backgroundColor="#FFFFFF"
} 
function mudar_cor_blur_impar(celula){
   celula.style.backgroundColor="#FFFFE1"
} 
function mudar_cor_blur_erro(celula){
   celula.style.backgroundColor="#CC0000"
}  
function checksubmit()
{
//  if (document.nota.pt.value == "")
//  {    alert("Por favor digite um peso para os Testes!")
//    document.nota.pt.focus()
//    return false
//  }
//    if (document.nota.pp.value == "")
//  {    alert("Por favor digite um peso para as Provas!")
//    document.nota.pp.focus()
//    return false
//  }
  return true
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
document.all.divEtapa.innerHTML ="<select class=borda></select>"
document.all.divTurma.innerHTML = "<select class=borda></select>"
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
document.all.divTurma.innerHTML = "<select class=borda></select>"
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
</script>
<script language="javascript"> 
  
    function keyPressed(TB, e, max_right, max_bottom)  
    { 
        if (e.keyCode == 40 || e.keyCode == 13) { // arrow down 
            if (TB.split("c")[0] < max_bottom) 
            document.getElementById(eval(TB.split("c")[0] + '+1') + 'c' + TB.split("c")[1]).focus(); 
            if (TB.split("c")[0] == max_bottom) 
            document.getElementById(1 + 'c' + TB.split("c")[1]).focus();


        } 
  
        if (e.keyCode == 38) { // arrow up 
            if(TB.split("c")[0] > 1) 
            document.getElementById(eval(TB.split("c")[0] + '-1') + 'c' + TB.split("c")[1]).focus(); 
            if (TB.split("c")[0] == 1) 
            document.getElementById(max_bottom + 'c' + TB.split("c")[1]).focus(); 
		
        } 
  
        if (e.keyCode == 37) { // arrow left 
            if(TB.split("c")[1] > 1) 
            document.getElementById(TB.split("c")[0] + 'c' + eval(TB.split("c")[1] + '-1')).focus();             
            if (TB.split("c")[1] == 1) 
            document.getElementById(TB.split("c")[0] + 'c' + max_right).focus(); 

		}   
  
        if (e.keyCode == 39) { // arrow right 
            if(TB.split("c")[1] < max_right) 
            document.getElementById(TB.split("c")[0] + 'c' + eval(TB.split("c")[1] + '+1')).focus();  
            if (TB.split("c")[1] == max_right) 
            document.getElementById(TB.split("c")[0] + 'c' + 1).focus(); 

		}                  
    } 
  
</script> 
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" background="../../../../img/fundo.gif" marginheight="0">
<%IF imprime="1"then
else
 call cabecalho (nivel) 
 end if%>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
    <td height="10" valign="top" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> </td>
  </tr>
    <%

			Set Conecta= Server.CreateObject("ADODB.Connection") 
		ABRIRgn = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		Conecta.Open ABRIRgn

call GeraNomes("PORT",unidade,curso,co_etapa,Conecta)

no_materia= session("no_materia")
no_unidades= session("no_unidades")
no_grau= session("no_grau")
no_serie= session("no_serie")
tp=	session("tp")

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("m", now) 
data = dia &"/"& mes &"/"& ano
horario = hora & ":"& min
acesso_prof = session("acesso_prof")


ST_Per_1 = ""
ST_Per_2 = ""
ST_Per_3 = ""

tp = session("tp")
%>
            <%if opt = "ok" then%>
            <tr>         
    <td height="10" valign="top"> 
      <%
		call mensagens(nivel,642,2,0)
%>
      <div align="center"></div></td>
            </tr>			
            <%elseif opt= "err6" then %>
            <tr> 
    <td height="10" valign="top"> 
      <%
	call mensagens(nivel,620,1,errante)
%>
</td>
            </tr>
            <%end if
%>
            <% IF trava="s" or (co_usr<>coordenador AND grupo="COO") then%>
            <tr>     
    <td height="10" valign="top"> 
      <%	call mensagens(nivel,9701,0,0) %>
</td>
            </tr>
			<% ELSEIF (autoriza=5 OR co_usr=coordenador) AND trava<>"s" AND ((periodo = 1 and ST_Per_1="x") OR (periodo = 2 and ST_Per_2="x") OR (periodo = 3 and ST_Per_3="x")) then%>
            <tr>     
    <td height="10" valign="top"> 
      <%	call mensagens(nivel,640,1,0) %>
</td>
            </tr>


            <%elseif (periodo = 1 and ST_Per_1="x") OR (periodo = 2 and ST_Per_2="x") OR (periodo = 3 and ST_Per_3="x") then%>
            <tr> 
    <td height="10" valign="top"> 
      <%
		call mensagens(nivel,624,0,0)
%>
</td>
            </tr>

            <% end if%>
<%if opt= "cln" then %>
            <tr> 
    <td height="10" valign="top"> 
      <%
	call mensagens(nivel,621,0,0)
%>
</td>
            </tr>
            <% end if%>						
	            <tr> 
    <td height="10" valign="top"> 
      <%	call mensagens(nivel,636,0,0) 

%>
    </td>
            </tr>			
            <tr class="tb_tit"> 
              
    <td height="15" class="tb_tit">&nbsp;Grade de Aulas</td>
            </tr>
            <tr> 
    <td height="36" valign="top"> 
<form name="alteracao" method="post" action="altera.asp">	
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="250" class="tb_subtit"> 
            <div align="center"><strong>UNIDADE 
              </strong></div></td>
          <td width="250" class="tb_subtit"> 
            <div align="center"><strong>CURSO 
              </strong></div></td>
          <td width="250" class="tb_subtit"> 
            <div align="center"><strong>ETAPA 
              </strong></div></td>
          <td width="250" class="tb_subtit"> 
            <div align="center"><strong>TURMA 
              </strong></div></td>
        </tr>
        <tr> 
                  <td> <div align="center"> 
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
                  <td> <div align="center"> 
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
                  <td> <div align="center"> 
                      <div id="divEtapa"> 
                        <select name="etapa" class="select_style" onChange="recuperarTurma(this.value)">
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
      </table>
</form>	  
	  </td>
            </tr>
      <tr> 
        
    <td valign="top"> 
      <%
	  
Call contalunos(CAMINHO_al,ano_letivo,unidade,curso,etapa,turma,"C")

co_usr=co_usr*1
coordenador=coordenador*1
autoriza=autoriza*1

if opt="cln" OR (periodo = 1 and ST_Per_1="x") OR (periodo = 2 and ST_Per_2="x") OR (periodo = 3 and ST_Per_3="x") OR trava="s" then
if nota ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na
Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,1,0)
else
	if nota="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb
	Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,2,0)
	else
		if nota ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc
		Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,3,0)
		else
			if nota ="TB_NOTA_E" then
			CAMINHOn = CAMINHO_ne
			Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,5,0)
			else
				if nota ="TB_NOTA_F" then
				CAMINHOn = CAMINHO_nf
				Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,6,0)
				else
					if nota ="TB_NOTA_K" then
					CAMINHOn = CAMINHO_nk			
					Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,7,0)
					else
						if nota ="TB_NOTA_L" then
						CAMINHOn = CAMINHO_nl			
						Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,8,0)
						else
							if nota ="TB_NOTA_M" then
							CAMINHOn = CAMINHO_nm			
							Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,9,0)
							else
							response.Write("ERRO TB_NOTA " &nota )
							End if
						End if
					End if
				End if
			End if
		End if
  	end if
end if

ELSEIF co_usr=coordenador AND trava<>"s"then


if nota ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na
Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,11,0)
else
	if nota="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb
	Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,21,0)
	else
		if nota ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc
		Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,31,0)
		else
			if nota ="TB_NOTA_E" then
			CAMINHOn = CAMINHO_ne
			Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,51,0)
			else
				if nota ="TB_NOTA_F" then
				CAMINHOn = CAMINHO_nf
				Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,61,0)
				else
					if nota ="TB_NOTA_K" then
					CAMINHOn = CAMINHO_nk			
					Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,71,0)
					else
						if nota ="TB_NOTA_L" then
						CAMINHOn = CAMINHO_nl			
						Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,81,0)
						else
							if nota ="TB_NOTA_M" then
							CAMINHOn = CAMINHO_nm			
							Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,91,0)
							else
							response.Write("ERRO TB_NOTA " &nota )
							End if
						End if
					End if
				End if
			End if
		End if
  	end if
end if
ELSEIF autoriza=5 and grupo<>"COO" AND trava<>"s"then

if nota ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na
Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,11,0)
else
	if nota="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb
	Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,21,0)
	else
		if nota ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc
		Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,31,0)
		else
			if nota ="TB_NOTA_E" then
			CAMINHOn = CAMINHO_ne
			Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,51,0)
			else
				if nota ="TB_NOTA_F" then
				CAMINHOn = CAMINHO_nf
				Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,61,0)
				else
					if nota ="TB_NOTA_K" then
					CAMINHOn = CAMINHO_nk			
					Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,71,0)
					else
						if nota ="TB_NOTA_L" then
						CAMINHOn = CAMINHO_nl			
						Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,81,0)
						else
							if nota ="TB_NOTA_M" then
							CAMINHOn = CAMINHO_nm			
							Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,91,0)
							else
							response.Write("ERRO TB_NOTA " &nota )
							End if
						End if
					End if
				End if
			End if
		End if
  	end if
end if
else
if nota ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na
Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,11,0)
else
	if nota="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb
	Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,21,0)
	else
		if nota ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc
		Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,31,0)
		else
			if nota ="TB_NOTA_E" then
			CAMINHOn = CAMINHO_ne
			Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,51,0)
			else
				if nota ="TB_NOTA_F" then
				CAMINHOn = CAMINHO_nf
				Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,61,0)
				else
					if nota ="TB_NOTA_K" then
					CAMINHOn = CAMINHO_nk			
					Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,71,0)
					else
						if nota ="TB_NOTA_L" then
						CAMINHOn = CAMINHO_nl			
						Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,81,0)
						else
							if nota ="TB_NOTA_M" then
							CAMINHOn = CAMINHO_nm			
							Call faltas(CAMINHOa,CAMINHOn,unidade,curso,co_etapa,turma,co_materia,periodo,ano_letivo,co_usr,91,0)
							else
							response.Write("ERRO TB_NOTA " &nota )
							End if
						End if
					End if
				End if
			End if
		End if
  	end if
end if
end if		
 %>
    </td>
      </tr>
  <%	
    Set RS = Nothing
	    Set RS2 = Nothing
		    Set RS3 = Nothing
End if
%>
      <tr>      
    <td height="40" valign="top"> <img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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