<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->



<!--#include file="../../../../inc/caminhos.asp"-->


<% 

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
cod= request.QueryString("cod_cons")
opt = request.QueryString("opt")
	

obr=cod


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")		

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
					
	
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod
		RS.Open SQL, CON1
		
nome_prof = RS("NO_Aluno")


		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod
		RS.Open SQL, CON1


ano_aluno = RS("NU_Ano")
situacao = RS("CO_Situacao")
encerramento= RS("DA_Encerramento")
unidade= RS("NU_Unidade")
curso= RS("CO_Curso")
etapa= RS("CO_Etapa")
turma= RS("CO_Turma")
cham= RS("NU_Chamada")
motivo= RS("DS_Motivo")



if situacao="C" then
remaneja="s"
else
remaneja="n"
end if

Call LimpaVetor2

call GeraNomes("PORT",unidade,curso,etapa,CON0)
no_unidade = session("no_unidades")
no_curso = session("no_grau")
no_etapa = session("no_serie")

			ano = DatePart("yyyy", now)
			mes = DatePart("m", now) 
			dia = DatePart("d", now) 
			hora = DatePart("h", now) 
			min = DatePart("n", now) 

 			if situacao="L" or situacao="E" or situacao="R" then
			data_exibe=encerramento			
			else
			data_exibe=dia&"/"&mes&"/"&ano
			end if

%>
<html>
<head>
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../../../js/mm_menu.js"></script>
<script type="text/javascript" src="../../../../js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
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

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}

function MM_showHideLayers() { //v6.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
    obj.visibility=v; }
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--

function checksubmit()
{
  if (document.busca.motivo.value == "")
  {    alert("Por favor digite um motivo para o remanejamento!")
    document.busca.motivo.focus()
    return false
  }	
	
var answer = confirm ("Tem certeza que deseja transferir o aluno <%response.Write(cod)%> de turma?")
if (!answer)
    return false
else
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
//-->
</script>
<script>
<!--
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
                                               oHTTPRequest.open("post", "executa.asp?ori=alt&opt=c", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
document.all.divEtapa.innerHTML ="<select name=etapa class=borda><option value=999990 selected></option></select>"
document.all.divTurma.innerHTML = "<select name=turma class=borda><option value=999990 selected></option></select>"
document.all.divChamada.innerHTML = "<select name=chamada class=borda><option value=999990 selected></option></select>"
                                                           }
                                               }
                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?ori=alt&opt=e", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select name=turma class=borda><option value=999990 selected></option></select>"
document.all.divChamada.innerHTML = "<select name=chamada class=borda><option value=999990 selected></option></select>"

                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?ori=alt&opt=t", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t
document.all.divChamada.innerHTML = "<select name=chamada class=borda><option value=999990 selected></option></select>"																	   
                                                           }
                                               }
                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }    
						 function recuperarChamada(tTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?ori=alt&opt=ch", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                       var resultado_ch= oHTTPRequest.responseText;
resultado_ch = resultado_ch.replace(/\+/g," ")
resultado_ch = unescape(resultado_ch)
document.all.divChamada.innerHTML = resultado_ch																	   
                                                           }
                                               }
                                               oHTTPRequest.send("t_pub=" + tTipo);											   
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
 <%if opt="ok" then%>
             <tr> 
         
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,9705,2,0) %>
    </td>
			  </tr>
 <%end if%>		 
  <%if remaneja="n" then%>
             <tr> 
         
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,416,1,0) %>
    </td>
			  </tr>
 <%else%> 	  
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,413,0,0) %>
    </td>
			  </tr>
 <%end if%> 			  			  
        <form action="bd.asp" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
          <tr>
      <td valign="top"> 
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo" height="206"
>
          <tr> 
            <td width="1000" class="tb_tit" height="15"
>Dados para Remanejamento</td>
            <td width="1000" class="tb_tit" height="15"
> </td>
          </tr>
          <tr> 
            <td height="21" width="1000"> <table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="19%" height="10"> <div align="right"><font class="form_dado_texto"> 
                      Matr&iacute;cula: </font></div></td>
                  <td width="9%" height="10"><font class="form_dado_texto"> 
                    <input name="cod" type="hidden" value="<%=cod%>">
                    <%response.Write(codigo)%>
                    </font></td>
                  <td width="6%" height="10"> <div align="right"><font class="form_dado_texto"> 
                      Nome: </font></div></td>
                  <td width="66%" height="10"><font class="form_dado_texto"> 
                    <%response.Write(nome_prof)%>
                    </font></td>
                </tr>
              </table></td>
            <td valign="top" width="1000" height="21"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font> </td>
          </tr>
          <tr> 
            <td height="19" bgcolor="#FFFFFF" width="1000">&nbsp;</td>
            <td valign="top" bgcolor="#FFFFFF" width="1000" height="19">&nbsp;</td>
          </tr>
          <tr> 
            <td class="tb_tit" height="15"
>Sai de:</td>
            <td class="tb_tit" height="15"
> </td>
          </tr>
          <tr> 
            <td colspan="2" width="1000" height="32"> <table width="100%" border="0" cellspacing="0">
                <tr class="tb_subtit"> 
                  <td width="200" height="10"> <div align="center">Unidade</div></td>
                  <td width="200" height="10"> <div align="center">Curso</div></td>
                  <td width="200" height="10"> <div align="center"> Etapa</div></td>
                  <td width="200" height="10"> <div align="center">Turma </div></td>
                  <td width="200" height="10"> <div align="center">Chamada</div></td>
                </tr>
                <tr class="tb_corpo"> 
                  <td width="200" height="10"> <div align="center"> <font class="form_dado_texto"> 
                    <input name="unidade_veio" type="hidden" value="<%=unidade%>">
                      <%response.Write(no_unidade)%>
                      </font></div></td>
                  <td width="200" height="10"> <div align="center"> <font class="form_dado_texto"> 
                    <input name="curso_veio" type="hidden" value="<%=curso%>">
                      <%response.Write(no_curso)%>
                      </font></div></td>
                  <td width="200" height="10"> <div align="center"> <font class="form_dado_texto"> 
                    <input name="etapa_veio" type="hidden" value="<%=etapa%>">
                      <%response.Write(no_etapa)%>
                      </font></div></td>
                  <td width="200" height="10"> <div align="center"> <font class="form_dado_texto"> 
                    <input name="turma_veio" type="hidden" value="<%=turma%>">
                      <%response.Write(turma)%>
                      </font></div></td>
                  <td width="200" height="10"> <div align="center"> <font class="form_dado_texto"> 
                    <input name="cham_veio" type="hidden" value="<%=cham%>">
                      <%response.Write(cham)%>
                      </font></div></td>
                </tr>
            </table></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" width="1000" height="19">&nbsp;</td>
            <td width="1000" height="19">&nbsp;</td>
          </tr>
          <tr> 
            <td class="tb_tit" height="15"
>Vai Para:</td>
            <td class="tb_tit" height="15"
> </td>
          </tr>
          <tr valign="top"> 
            <td width="1000" height="48" colspan="2">
  <%if remaneja="s" then%>			
			<table width="100%" border="0" cellspacing="0">
                <tr class="tb_subtit"> 
                  <td width="200" height="10"> <div align="center">Unidade</div></td>
                  <td width="200" height="10"> <div align="center">Curso</div></td>
                  <td width="200" height="10"> <div align="center"> Etapa</div></td>
                  <td width="200" height="10"> <div align="center">Turma </div></td>
                  <td width="200" height="10"> <div align="center">Chamada</div></td>
                </tr>
                <tr valign="top" class="tb_corpo"> 
                  <td width="200" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <select name="unidade" class="borda" onChange="recuperarCurso(this.value)">
                        <%		

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0
NU_Unidade_Check=999999		
While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")
if NU_Unidade = NU_Unidade_Check then
RS0.MOVENEXT		
unidade=unidade*1
NU_Unidade=NU_Unidade*1		
elseif NU_Unidade = unidade then
%>
                        <option value="<%response.Write(NU_Unidade)%>" selected> 
                        <%response.Write(NO_Abr)%>
                        </option>
                        <%
NU_Unidade_Check = NU_Unidade
RS0.MOVENEXT							  
else%>
                        <option value="<%response.Write(NU_Unidade)%>"> 
                        <%response.Write(NO_Abr)%>
                        </option>
                        <%

NU_Unidade_Check = NU_Unidade
RS0.MOVENEXT
end if
WEND
%>
                      </select>
                      </font></div></td>
                  <td width="200" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <div id="divCurso"><select name="curso" class="borda" onChange="recuperarEtapa(this.value)">
                        <%		

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT Distinct CO_Curso FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade
		RS0.Open SQL0, CON0
		
While not RS0.EOF
CO_Curso = RS0("CO_Curso")

		Set RS0a = Server.CreateObject("ADODB.Recordset")
		SQL0a = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
		RS0a.Open SQL0a, CON0
		
NO_Curso = RS0a("NO_Curso")

if CO_Curso = curso then
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
</select></div>
                      </font></div></td>
                  <td width="200" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <div id="divEtapa">
<select name="etapa" class="borda" onChange="recuperarTurma(this.value)">
                        <%		

		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"'"
		RS0b.Open SQL0b, CON0
		
While not RS0b.EOF
CO_Etapa = RS0b("CO_Etapa")

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&curso&"' AND CO_Etapa='"&CO_Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")
if CO_Etapa = etapa then
%>
                        <option value="<%response.Write(CO_Etapa)%>" selected> 
                        <%response.Write(NO_Etapa)%>
                        </option>
                        <%
else								
%>
                        <option value="<%response.Write(CO_Etapa)%>"> 
                        <%response.Write(NO_Etapa)%>
                        </option>
                        <%
end if
RS0b.MOVENEXT
WEND
%>
</select>                      
                      
                      </div>
                      </font></div></td>
                  <td width="200" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <div id="divTurma">
<select name="turma" class="borda" onChange="recuperarChamada(this.value)">
    <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Turma where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"' AND CO_Etapa='" & etapa & "' order by CO_Turma" 
		RS3.Open SQL3, CON0						

while not RS3.EOF
co_turma= RS3("CO_Turma")
'cap_turma= RS3("NU_Capacidade")
'
'		Set RS_al = Server.CreateObject("ADODB.Recordset")
'		SQL_al = "SELECT COUNT(CO_Matricula) AS alunos_turma FROM TB_Matriculas where NU_Ano="&ano_letivo&" AND NU_Unidade="&unidade&" AND CO_Curso='"&curso&"' AND CO_Etapa='" & etapa & "'  AND CO_Turma='" & co_turma & "'" 
'		RS_al.Open SQL_al, CON1
'								
'alunos_turma=RS_al("alunos_turma")
'cap_turma=cap_turma*1
'alunos_turma=alunos_turma*1
'vagas_turma=cap_turma-alunos_turma
'texto_turma= co_turma&" - "&vagas_turma
texto_turma= co_turma

if co_turma = turma then
 %>
<option value="<%=response.Write(co_turma)%>" selected> 
    <%response.Write(texto_turma)%>
    </option> 
    <%
else
 %>
<option value="<%=response.Write(co_turma)%>"> 
    <%response.Write(texto_turma)%>
    </option> 
    <%
end if
RS3.MOVENEXT
WEND
%></select>                      
                      </div>
                      </font></div></td>
                  <td width="200" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <div id="divChamada">
                   <select name="chamada" id="chamada" class="borda"  > 					 
<%nu_chamada_ckq=0
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Matriculas where NU_Unidade="&unidade&" AND CO_Curso='"&curso&"' AND CO_Etapa='" & etapa & "' AND CO_Turma='" & turma& "' AND NU_Ano="&ano_letivo&" order by NU_Chamada"
		RS4.Open SQL4, CON1
		
while not RS4.EOF
nu_chamada=RS4("NU_Chamada")
cham=cham*1
nu_chamada=nu_chamada*1
	if (nu_chamada_ckq <>nu_chamada - 1) then
		teste_nu_chamada = nu_chamada-nu_chamada_ckq
		teste_nu_chamada=teste_nu_chamada-1
		for k=1 to teste_nu_chamada 
			nu_chamada_falta=nu_chamada_ckq+1
				 %>	
                <option value="<%=response.Write(nu_chamada_falta)%>"> 
				<%response.Write(nu_chamada_falta)%>
				</option>  
				<%
			nu_chamada_ckq=nu_chamada_falta
		next
	end if
nu_chamada_ckq=nu_chamada	
RS4.MOVENEXT	
wend
if RS4.EOF then
nu_chamada=nu_chamada*1
ultima_chamada=nu_chamada+1
	%>
	 <option value="<%=response.Write(ultima_chamada)%>" selected> 
	<%response.Write(ultima_chamada)%>
	</option>  	
<%end if	
 %>				  
		   </select> 
                    
                      </div>
                      </font></div></td>
                </tr>
                <tr valign="top" class="tb_subtit">
                  <td width="200" height="10"><div align="center">Data do Remanejamento</div></td>
                  <td width="200" height="10">&nbsp;&nbsp;&nbsp;Motivo do Remanejamento</td>
                  <td width="200" height="10">&nbsp;</td>
                  <td width="200" height="10">&nbsp;</td>
                  <td width="200" height="10">&nbsp;</td>
                </tr>
<%
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
data = dia &"/"& mes &"/"& ano
data_motivo_dividida = split(data,"/")		
%>                
                <tr valign="top" class="tb_corpo">
                  <td width="200" height="10"><select name="dia_remaneja" class="textInput" id="dia_remaneja">
					<%for d=1 to 31
					if d<10 then
						dia="0"&d
					else
						dia=d
					end if	
					data_motivo_dividida(0)=data_motivo_dividida(0)*1
					d=d*1
						if data_motivo_dividida(0)=d then				
					%>
                        <option value="<%response.Write(dia)%>" selected><%response.Write(dia)%></option> 
                    	<%else%> 
                        <option value="<%response.Write(dia)%>"><%response.Write(dia)%></option>  						<%end if%>                          
					<%next%>                                      
		            </select>/
		          <select name="mes_remaneja" class="textInput" id="mes_remaneja">
					<%
					data_motivo_dividida(1)=data_motivo_dividida(1)*1
					if data_motivo_dividida(1)=1 then%>	                    
                     <option value="1" selected>Janeiro</option> 
                    <%else%>  
                     <option value="1" >Janeiro</option>
                    <%end if
					if data_motivo_dividida(1)=2 then%>	                    
                     <option value="2" selected>Fevereiro</option> 
                    <%else%>  
                     <option value="2" >Fevereiro</option>
                    <%end if
					if data_motivo_dividida(1)=3 then%>	                    
                     <option value="3" selected>Março</option> 
                    <%else%>  
                     <option value="3" >Março</option>
                    <%end if	
					if data_motivo_dividida(1)=4 then%>	                    
                     <option value="4" selected>Abril</option> 
                    <%else%>  
                     <option value="4" >Abril</option>
                    <%end if
					if data_motivo_dividida(1)=5 then%>	                    
                     <option value="5" selected>Maio</option> 
                    <%else%>  
                     <option value="5" >Maio</option>
                    <%end if																					
					if data_motivo_dividida(1)=6 then%>	                    
                     <option value="6" selected>Junho</option> 
                    <%else%>  
                     <option value="6" >Junho</option>
                    <%end if
					if data_motivo_dividida(1)=7 then%>	                    
                     <option value="7" selected>Julho</option> 
                    <%else%>  
                     <option value="7" >Julho</option>
                    <%end if
					if data_motivo_dividida(1)=8 then%>	                    
                     <option value="8" selected>Agosto</option> 
                    <%else%>  
                     <option value="8" >Agosto</option>
                    <%end if
					if data_motivo_dividida(1)=9 then%>	                    
                     <option value="9" selected>Setembro</option> 
                    <%else%>  
                     <option value="9" >Setembro</option>
                    <%end if															
					if data_motivo_dividida(1)=10 then%>	                    
                     <option value="10" selected>Outubro</option> 
                    <%else%>  
                     <option value="10" >Outubro</option>
                    <%end if	
					if data_motivo_dividida(1)=11 then%>	                    
                     <option value="11" selected>Novembro</option> 
                    <%else%>  
                     <option value="11" >Novembro</option>
                    <%end if	
					if data_motivo_dividida(1)=12 then%>	                    
                     <option value="12" selected>Dezembro</option> 
                    <%else%>  
                     <option value="12" >Dezembro</option>
                    <%end if%>
                    </select>/
		          <select name="ano_remaneja" class="textInput" id="ano_remaneja">
					<%for a=data_motivo_dividida(2) to data_motivo_dividida(2)+1				
					data_motivo_dividida(2)=data_motivo_dividida(2)*1
					a=a*1
						if data_motivo_dividida(2)=a then				
					%>
                        <option value="<%response.Write(a)%>" selected><%response.Write(a)%></option>
						<%else%> 
                        <option value="<%response.Write(a)%>"><%response.Write(a)%></option>  						
						<%end if%>                            
					<%next%>                                      
	              </select></td>
                  <td height="10" colspan="4">&nbsp;&nbsp;&nbsp;<input name="motivo" type="text" id="motivo" class="borda" size="100" maxlength="255"></td>
                </tr>
              </table>
<%end if%>
			  </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td height="22" colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr class="tb_corpo"> 
                  <td colspan="3"><hr></td>
                </tr>
                <tr> 
                  <td width="33%"><div align="center"> 
                      <input type="button" name="Submit2" value="Voltar" class="borda_bot3" onClick="MM_goToURL('parent','index.asp?nvg=WS-MA-MA-REA')">
                    </div></td>
                  <td width="34%">&nbsp;</td>
                  <td width="33%"> <div align="center"> 
				    <%if remaneja="s" then%>
                      <input type="submit" name="Submit" value="Confirmar" class="borda_bot">
					 <%end if%> 
                    </div></td>
                </tr>
              </table>
              
            </td>
          </tr>
        </table></td>
    </tr>
</form>
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

</body>
<script type="text/javascript">
<!--
//  initInputHighlightScript();
//-->
</script>


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