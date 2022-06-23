<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes_comuns.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes7.asp"-->
<!--#include file="../../../../inc/graficos.asp"-->
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

'if opt= "vt" then
dados= split(obr, "_" )
unidade= dados(0)
curso= dados(1)
co_etapa= dados(2)
tipo_parcela = dados(3)


'else
'unidade = request.Form("unidade")
'curso = request.Form("curso")
'co_etapa = request.Form("etapa")
'tipo_parcela = request.Form("tipo")
'end if

ano_letivo = session("ano_letivo")
obr=unidade&"_"&curso&"_"&co_etapa&"_"&tipo_parcela&"_"&ano_letivo&"_"&opt


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

		Set CON7 = Server.CreateObject("ADODB.Connection") 
		ABRIR7 = "DBQ="& CAMINHO_pf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON7.Open ABRIR7			


if unidade="nulo" then
	no_unidade="sem unidade"
	sql_unidade = ""	
	sql_curso = ""
	sql_etapa = ""	
else	
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
	no_unidade = RS0("NO_Unidade")
	sql_unidade =" AND NU_Unidade="& unidade

	if isnull(curso) or curso="" or curso="999990" then
		no_etapa="sem curso"
		sql_curso = ""
		sql_etapa = ""	
	else
		if curso=999990 then
			no_etapa="sem curso"
			sql_curso = ""	
			sql_etapa = ""				
		else
			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
			RS1.Open SQL1, CON0
			sql_curso = " AND CO_Curso ='"& curso &"'"		
			no_curso = RS1("NO_Abreviado_Curso")
	
			if isnull(co_etapa) or co_etapa="" or co_etapa="999990" then
				no_etapa="sem etapa"	
				sql_etapa = ""	
			else
				if isnumeric(co_etapa) then
					if co_etapa=999990 then
						no_etapa="sem etapa"	
						sql_etapa = ""				
					else
						Set RS3 = Server.CreateObject("ADODB.Recordset")
						SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' and CO_Curso ='"& curso &"'"  
						RS3.Open SQL3, CON0
									
						if RS3.EOF THEN
							no_etapa="sem etapa"
						else
							no_etapa=RS3("NO_Etapa")
							sql_etapa = " AND CO_Etapa ='"& co_etapa &"'"	
						end if
					end if	
				else
					no_etapa="sem etapa"	
					sql_etapa = ""				
				end if	
			end if	
		end if	
	end if
end if	

call navegacao (CON,chave,nivel)
navega=Session("caminho")
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
						
						
						 function recuperarOrder2(oTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=o2", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                    var resultado_o  = oHTTPRequest.responseText;
resultado_o = resultado_o.replace(/\+/g," ")
resultado_o = unescape(resultado_o)
document.all.divOrder2.innerHTML =resultado_o
//recuperarEtapa()
                                                           }
                                               }

                                               oHTTPRequest.send("order_pub=" + oTipo);
                                   }


						 function recuperarOrder3(oTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=o3", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                    var resultado_o  = oHTTPRequest.responseText;
resultado_o = resultado_o.replace(/\+/g," ")
resultado_o = unescape(resultado_o)
document.all.divOrder3.innerHTML =resultado_o
//recuperarEtapa()
                                                           }
                                               }

                                               oHTTPRequest.send("order_pub=" + oTipo);
                                   }
								   
						 function recuperarOrder4(oTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=o4", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                    var resultado_o  = oHTTPRequest.responseText;
resultado_o = resultado_o.replace(/\+/g," ")
resultado_o = unescape(resultado_o)
document.all.divOrder4.innerHTML =resultado_o
//recuperarEtapa()
                                                           }
                                               }

                                               oHTTPRequest.send("order_pub=" + oTipo);
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
	call mensagens(nivel,913,0,0) 
%>
    </td>
                </tr>
                <tr> 
                  
    <td valign="top"> 
      <form name="inclusao" method="post" action="../../../../relatorios/swd030.asp">
                <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
                  <tr class="tb_tit"
> 
                    <td width="653" height="15" class="tb_tit"
>Gerar Relatório<!--Ordena&ccedil;&atilde;o--><input name="obr" type="hidden" value="<%response.write(obr)%>"></td>
                  </tr>
                  <tr> 
                    
            <td><table width="998" border="0" cellspacing="0">
<!--                <tr> 
                  <td width="25%" class="tb_subtit"> <div align="center">PRIMEIRO CRIT&Eacute;RIO</div></td>
                  <td width="25%" class="tb_subtit"> <div align="center">SEGUNDO CRIT&Eacute;RIO</div></td>
                  <td width="25%" class="tb_subtit"> <div align="center">TERCEIRO CRIT&Eacute;RIO</div></td>
                  <td width="25%" class="tb_subtit"><div align="center">QUARTO CRIT&Eacute;RIO</div></td>
                  </tr>
                <tr>
                  <td background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF"> 
                    <div align="center"> 
                      <select name="order1" class="select_style" onChange="recuperarOrder2(this.value)">
                        <option value="nulo" selected>
                        </option>
                        <option value="mes"> Mês em Aberto
                        </option>							
                        <option value="UCET"> U/C/E/T
                        </option>
                        <option value="matricula"> Matricula
                        </option>						
                        <option value="nome"> Nome do Aluno
                        </option>		
                        <option value="lancamento"> Tipo de Lançamento
                        </option>																	
                      </select>
                    </div></td>
                  <td background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF"> 
                    <div align="center"> 
                      <div id="divOrder2"> 
                        <select name="order2" class="select_style">
                        </select>
                      </div>
                    </div></td>
                  <td background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF"> 
                  	<div align="center"> 
                  		<div id="divOrder3"> 
                        <select name="order3" class="select_style">
                        </select>                  			</div>
                  		</div></td>
                  <td width="25%"><div align="center">
                  		<div id="divOrder4"> 
                        <select name="order4" class="select_style">
                        </select>                  			</div>
                  	</div></td>
                  </tr>
                <tr>
                	<td colspan="4" bgcolor="#FFFFFF"><hr></td>
                	</tr>-->
                <tr>
                	<td height="15" colspan="4" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                		<tr>
                			<td width="33%" height="15" bgcolor="#FFFFFF"></td>
                			<td width="34%" height="15" bgcolor="#FFFFFF"></td>
                			<td width="33%" height="15" align="center" bgcolor="#FFFFFF"><font size="2" face="Arial, Helvetica, sans-serif">
                				<input name="Submit4" type="submit" class="botao_prosseguir" id="Submit4" value="Procurar">
                				</font></td>
                			</tr>
                		</table></td>
                	</tr>
              </table></td>
                  </tr>
                  <tr> 
                    
            <td align="center" valign="top">&nbsp;</td>
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