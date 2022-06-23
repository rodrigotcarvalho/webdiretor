<%On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<% 
cod= request.form("cod")	
co_mat_fil= request.form("mat_prin")
periodo= request.form("periodo")
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
ori = request.QueryString("or")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
	
obr=cod&"_"&co_mat_fil&"_"&periodo


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
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CONa = Server.CreateObject("ADODB.Connection") 
		ABRIRa = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONa.Open ABRIRa		

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod
		RS.Open SQL, CON1
		
		
codigo = RS("CO_Matricula")
nome_prof = RS("NO_Aluno")



		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod
		RS.Open SQL, CON1


ano_aluno = RS("NU_Ano")
rematricula = RS("DA_Rematricula")
situacao = RS("CO_Situacao")
encerramento= RS("DA_Encerramento")
unidade= RS("NU_Unidade")
curso= RS("CO_Curso")
etapa= RS("CO_Etapa")
turma= RS("CO_Turma")
cham= RS("NU_Chamada")




Call LimpaVetor2

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
  if (document.busca.busca1.value != "" && document.busca.busca2.value != "")
  {    alert("Por favor digite SOMENTE uma opção de busca!")
    document.busca.busca1.focus()
    return false
  }
    if (document.busca.busca1.value == "" && document.busca.busca2.value == "")
  {    alert("Por favor digite uma opção de busca!")
    document.busca.busca1.focus()
    return false
  }
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
} function createXMLHTTP()
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
						
						

						 function recuperarPeriodo()
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

                                               oHTTPRequest.send();
                                   }
								   

 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
</head>
<% if opt="listall" or opt="list" then%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%else %>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%end if %>
<%call cabecalho(nivel)
%>
<div id="fundo" style="position:absolute; left:0px; top:0px; width:100%; height:100%; z-index:1; background-color: #000000; layer-background-color: #000000; border: 1px none #000000; visibility: hidden;" class="transparente"></div>
<div id="alinha" style="position:absolute; width:400px; visibility: hidden; z-index: 2; left: 326px; height: 520px;"> 
  <table width="300" border="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr> 
      <td width="478"> <div align="right"> <span class="voltar1"><font class="form_dado_texto"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')">fechar</a></font></span></div></td>
      <td width="20"><div align="right"><span class="voltar1"><font class="form_dado_texto"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')"><img src="../../../../img/fecha.gif" width="20" height="16" border="0"></a></font></span></div></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center" ><img src="../../../../img/fotos/aluno/<% =codigo %>.jpg" height="500"></div></td>
    </tr>
    <tr>
      <td colspan="2"><div align="center"><font class="form_dado_texto">
          <%response.Write(nome_prof)%>
          </font></div></td>
    </tr>
  </table>
</div>

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
      <%call mensagens(nivel,18,0,0) %>
    </td>
			  </tr>			  
<form name="form1" method="post" action="grafico.asp?ori=01">
          <tr>
      <td valign="top">
<table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
          <tr> 
            <td width="653" class="tb_tit"
>Dados Escolares</td>
            <td width="113" class="tb_tit"
> </td>
          </tr>
          <tr> 
            <td height="10"> <table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="19%" height="10"> <div align="right"><font class="form_dado_texto">Matr&iacute;cula: 
                      </font></div></td>
                  <td width="9%" height="10"> <font class="form_dado_texto"> 
                    <input name="cod" type="hidden" value="<%=codigo%>">
                    <%response.Write(codigo)%>
                    </font></td>
                  <td width="6%" height="10"> <div align="right"><font class="form_dado_texto">Nome: 
                      </font></div></td>
                  <td width="66%" height="10"> <font class="form_dado_texto"> 
                    <%response.Write(nome_prof)%>
                    <input name="nome" type="hidden" class="textInput" id="nome2"  value="<%response.Write(nome_prof)%>" size="75" maxlength="50">
                    &nbsp; </font></td>
                </tr>
              </table></td>
            <td valign="top">&nbsp; </td>
          </tr>
          <tr> 
            <td height="10" bgcolor="#FFFFFF">&nbsp;</td>
            <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="2"><table width="100%" border="0" cellspacing="0">
                <tr class="tb_subtit"> 
                  <td width="33" height="10"> <div align="center"> 
                      <%
call GeraNomes("PORT",unidade,curso,etapa,CON0)
no_unidades = session("no_unidades")
no_grau = session("no_grau")
no_serie = session("no_serie")
%>
                      Ano</div></td>
                  <td width="81" height="10"> <div align="center">Matr&iacute;cula</div></td>
                  <td width="75" height="10" class="tb_subtit"> <div align="center">Cancelamento</div></td>
                  <td width="86" height="10"> <div align="center"> Situa&ccedil;&atilde;o</div></td>
                  <td width="54" height="10"> <div align="center">Chamada</div></td>
                </tr>
                <tr class="tb_corpo"
> 
                  <td width="33" height="10"> <div align="center"><font class="form_dado_texto"> 
                      <%response.Write(ano_aluno)%>
                      </font> </div></td>
                  <td width="81" height="10"> <div align="center"><font class="form_dado_texto"> 
                      <%response.Write(rematricula)%>
                      </font></div></td>
                  <td width="75" height="10"> <div align="center"><font class="form_dado_texto"> 
                      <%response.Write(encerramento)%>
                      </font></div></td>
                  <td width="86" height="10"> <div align="center"><font class="form_dado_texto"> 
                      <%
					
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
		RSCONTST.Open SQLCONTST, CON0
							
				no_situacao = RSCONTST("TX_Descricao_Situacao")	
					response.Write(no_situacao)%>
                      </font></div></td>
                  <td width="54" height="10"> <div align="center"><font class="form_dado_texto"> 
                      <%response.Write(cham)%>
                      </font></div></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td colspan="2" bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="2"> <table width="100%" border="0" cellspacing="0">
                <tr class="tb_subtit"> 
                  <td width="150" height="10"> 
                    <div align="center"> 
                      Unidade</div></td>
                  <td width="200" height="10"> 
                    <div align="center"> Curso</div></td>
                  <td width="150" height="10"> 
                    <div align="center"> Etapa</div></td>
                  <td width="150" height="10"> 
                    <div align="center"> Turma </div></td>
                  <td width="200"> 
                    <div align="center">Disciplina</div></td>
                  <td width="150"> 
                    <div align="center">Per&iacute;odo</div></td>
                </tr>
                <tr> 
                  <td width="150" height="10" class="tb_corpo"
> 
                    <div align="center"><font class="form_dado_texto"> 
                      <%response.Write(no_unidades)%>
                      </font></div></td>
                  <td width="200" height="10" class="tb_corpo"
> 
                    <div align="center"><font class="form_dado_texto"> 
                      <%response.Write(no_grau)%>
                      </font></div></td>
                  <td width="150" height="10" class="tb_corpo"
> 
                    <div align="center"><font class="form_dado_texto"> 
                      <%response.Write(no_serie)%>
                      </font></div></td>
                  <td width="150" height="10" class="tb_corpo"
> 
                    <div align="center"><font class="form_dado_texto"> 
                      <%response.Write(turma)%>
                      </font></div></td>
                  <td width="200"> 
                    <div align="center">
				                        <select name="mat_prin" class="select_style" onChange="recuperarPeriodo();">
                        <%				  
			Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
		RS5.Open SQL5, CON0

while not RS5.EOF
co_mat_prin= RS5("CO_Materia")


		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Materia where CO_Materia ='"& co_mat_prin &"'"
		RS7.Open SQL7, CON0
		
		no_mat_prin= RS7("NO_Materia")
		
if co_mat_prin=co_mat_fil then
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
WEND			  

%>
                      </select>
                      </div></td>
                  <td width="150"> 
				<div id="divPeriodo" align="center"> 
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
 </select></div></td>
                </tr>
                <tr> 
                  <td height="10" colspan="6" class="tb_corpo"
>&nbsp;</td>
                </tr>
                <tr>
                  <td height="10" colspan="6" class="tb_corpo"
>
                    <%
		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Materia where CO_Materia ='"& co_mat_fil &"'"
		RS7.Open SQL7, CON0					
co_mat_prin= RS7("CO_Materia_Principal")

if isnull(co_mat_prin) or co_mat_prin="" then
co_mat_prin=co_mat_fil
end if			

		Set RSFIL = Server.CreateObject("ADODB.Recordset")
		SQLFIL = "SELECT * FROM TB_Da_Aula where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& etapa &"' AND CO_Turma ='"& turma &"' AND CO_Materia_Principal='"&co_mat_prin&"' AND CO_Materia='"&co_mat_fil&"'" 
		RSFIL.Open SQLFIL, CON2

	notaFIL=RSFIL("TP_Nota")
	co_mat_prin=RSFIL("CO_Materia")

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
	
		Set RSnFIL = Server.CreateObject("ADODB.Recordset")
		SQLnFIL = "SELECT * FROM "&notaFIL&" where CO_Matricula ="& codigo &" AND CO_Materia ='"& co_mat_fil &"'AND CO_Materia_Principal ='"& co_mat_prin &"' AND CO_Materia='"&co_mat_fil&"' And NU_Periodo="&periodo
		RSnFIL.Open SQLnFIL, CONn


if RSnFIL.eof then
h_al=1
else
m_al=RSnFIL("VA_Media3")

if m_al="" or isnull(m_al) then
h_al=1
else
h_al=m_al*3.225
end if
end if



		Set RSt0 = Server.CreateObject("ADODB.Recordset")
		SQLt0 = "SELECT * FROM TB_Aluno_Esta_Turma where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& etapa &"'"
		RSt0.Open SQLt0, CONa
m_disc=0
m_disca=0
div_disc=0		
while not RSt0.EOF
codigo0= RSt0("CO_Matricula")


		Set RSnFIL2 = Server.CreateObject("ADODB.Recordset")
		SQLnFIL2 = "SELECT * FROM "&notaFIL&" where CO_Matricula ="& codigo0 &" AND CO_Materia ='"& co_mat_fil &"'AND CO_Materia_Principal ='"& co_mat_prin &"' AND CO_Materia='"&co_mat_fil&"' And NU_Periodo="&periodo
		RSnFIL2.Open SQLnFIL2, CONn

'response.Write("SQLnFIL2 = SELECT * FROM "&notaFIL&" where CO_Materia ='"& co_mat_fil &"'AND CO_Materia_Principal ='"& co_mat_prin &"'AND NU_Periodo="&periodo)

if RSnFIL2.eof then
m_disca=m_disca
div_disc=div_disc
else
m_disc=RSnFIL2("VA_Media3")

if m_disc="" or isnull(m_disc) then
m_disca=m_disca
div_disc=div_disc
else
m_disc=m_disc*1
m_disca=m_disca+m_disc
div_disc=div_disc+1
end if
end if
RSt0.MOVENEXT
wend

if div_disc=0 then
h_d=1
else	

media_disc=m_disca/div_disc
media_disc=formatNumber(media_disc,0)
h_d=media_disc*3.225
end if

		Set RSt = Server.CreateObject("ADODB.Recordset")
		SQLt = "SELECT * FROM TB_Aluno_Esta_Turma where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& etapa &"' AND CO_Turma ='"& turma &"'"
		RSt.Open SQLt, CONa
m_t=0
m_ta=0
div_t=0		
while not RSt.EOF
codigo2= RSt("CO_Matricula")

		Set RSnFIL3 = Server.CreateObject("ADODB.Recordset")
		SQLnFIL3 = "SELECT * FROM "&notaFIL&" where CO_Matricula ="& codigo2 &" AND CO_Materia ='"& co_mat_fil &"'AND CO_Materia_Principal ='"& co_mat_prin &"' AND CO_Materia='"&co_mat_fil&"' And NU_Periodo="&periodo
		RSnFIL3.Open SQLnFIL3, CONn
	
	
if RSnFIL3.eof then
m_ta=m_ta
div_t=div_t
else	
m_t=RSnFIL3("VA_Media3")

if m_t="" or isnull(m_t)then
m_ta=m_ta
div_t=div_t
else
m_t=m_t*1
'response.write("->"&m_t)
m_ta=m_ta+m_t

div_t=div_t+1
'response.write("-a-"&m_ta&"-d"&div_t&"<br>")
end if
end if
RSt.MOVENEXT
wend
if div_t=0 then
h_t=1
else	
media_t=m_ta/div_t
media_t=formatNumber(media_t,0)
h_t=media_t*3.225
end if		
%>
                    <table width="538" height="387" border="0" align="center" cellspacing="0">
                      <tr> 
                        <td height="345" valign="bottom" background="../../../../img/grafico/fundo_nota.jpg"> 
                          <table width="300" height="340" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr valign="bottom"> 
                              <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../../../../img/grafico/3.gif" width="100" height="<%=h_al%>"></font></td>
                              <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../../../../img/grafico/2.gif" width="100" height="<%=h_d%>"></font></td>
                              <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../../../../img/grafico/1.gif" width="100" height="<%=h_t%>"></font></td>
                            </tr>
                            <tr> 
                              <td height="9" colspan="3"><img src="../../../../img/grafico/espaco_nota.gif" width="21" height="9"></td>
                            </tr>
                          </table></td>
                      </tr>
                      <tr> 
                        <td height="21"> <div align="center"> 
                            <%if periodo=1 then%>
                            <img src="../../../../img/grafico/p1.jpg" width="150" height="21"> 
                            <%elseif periodo=2 then%>
                            <img src="../../../../img/grafico/p2.jpg" width="150" height="21"> 
                            <%elseif periodo=3 then%>
                            <img src="../../../../img/grafico/p3.jpg" width="150" height="21"> 
                            <%elseif periodo=4 then%>
                            <img src="../../../../img/grafico/p4.jpg" width="150" height="21"> 
                            <%elseif periodo=5 then%>
                            <img src="../../../../img/grafico/p5.jpg" width="150" height="21"> 
                            <%elseif periodo=6 then%>
                            <img src="../../../../img/grafico/p6.jpg" width="150" height="21"> 														
                            <%end if%>
                          </div></td>
                      </tr>
                      <tr> 
                        <td height="21"><table width="410" border="0" align="center" cellspacing="0">
                            <tr> 
                              <td width="2%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../../../../img/grafico/3.gif" width="10" height="10"></font></td>
                              <td width="12%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Aluno</font></td>
                              <td width="86%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                                <%response.Write(m_al)%>
                                <img src="../../../../img/grafico/espaco_nota.gif" width="21" height="9"></font></td>
                            </tr>
                            <tr> 
                              <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../../../../img/grafico/2.gif" width="10" height="10"></font></td>
                              <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Etapa</font></td>
                              <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                                <%response.Write(media_disc)%>
                                <img src="../../../../img/grafico/espaco_nota.gif" width="21" height="9"></font></td>
                            </tr>
                            <tr> 
                              <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../../../../img/grafico/1.gif" width="10" height="10"></font></td>
                              <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Turma</font></td>
                              <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                                <%response.Write(media_t)%>
                                <img src="../../../../img/grafico/espaco_nota.gif" width="21" height="9"></font></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
          </tr>
        </table>
      </td>
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