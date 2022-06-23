<%On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->

<!--#include file="../../../../inc/funcoes2.asp"-->

<%
session("nvg")=""
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
nvg=request.QueryString("nvg")
opt = request.QueryString("opt")
chave=nvg
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
Session("data_consulta")=""
Session("hora_consulta")=""


		
		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2	

 call VerificaAcesso (CON,chave,nivel)
autoriza=Session("autoriza")

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
}  function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
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
document.all.divEtapa.innerHTML ="<select name='etapa' class='select_style' id='etapa'><option value='999990' selected>           </option></select>"
document.all.divTurma.innerHTML = "<select name='turma' class='select_style' id='turma'><option value='999990' selected>           </option></select>"
document.all.divDisciplina.innerHTML = "<select name='mat_prin' class='select_style' id='periodo'><option value='999999' selected>           </option></select>"
//recuperarEtapa()
                                                           }
                                               }
 
                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }
 
 
						 function recuperarEtapa(cTipo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=e10", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select name='turma' class='select_style' id='turma'><option value='999990' selected>           </option></select>"
document.all.divDisciplina.innerHTML = "<select name='mat_prin' class='select_style' id='periodo'><option value='999999' selected>           </option></select>"
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
																	   
                                                           }
                                               }
 
                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
function recuperarPeriodo(eTipo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=p1", true);
 
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
function recuperarDisciplina(cTipo, eTipo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=d5", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                       var resultado_d= oHTTPRequest.responseText;
resultado_d = resultado_d.replace(/\+/g," ")
resultado_d = unescape(resultado_d)
document.all.divDisciplina.innerHTML = resultado_d
																	   
                                                           }
                                               }
 
                                               oHTTPRequest.send("c_pub=" + cTipo +"&e_pub=" + eTipo);
                                   }									   
//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif"leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<%call cabecalho(nivel)%>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
                  <tr>                    
            
    <td height="10" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> </td>
  </tr>
  <tr>                   
    <td height="10"> 
      <%
	  if autoriza="no" then
	  	call mensagens(4,9700,1,0) 	  
	  else
	  	call mensagens(4,617,0,0) 
	  end if%>
    </td>
                  </tr>				  
		  				  				  

  <tr> 
    <td valign="top">

<table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">

          <tr> 
            <td> 
              <%	  if autoriza="no" then			
		else
ano_slct = DatePart("yyyy", now)
mes_slct = DatePart("m", now) 
dia_slct = DatePart("d", now) 
ano_slct=ano_slct*1
ano_letivo=ano_letivo*1
if ano_slct=ano_letivo then
	mes_compara=mes_slct
	dia_compara=dia_slct	
else
	mes_compara=12
	dia_compara=31	
end if										

%>
        
      <table width="1000" border="0" cellspacing="0">
        <tr> 
                <td valign="top"> 
                  <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo"
>
                    <tr> 
                      <td class="tb_tit">Período de Monitoramento de Notas</td>
                </tr>
                <tr> 
                  <td> <form name="form1" method="post" action="monitora.asp?opt=1&nvg=<%=nvg%>">
                          <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0">
                            <tr>
                            	<td width="100%"><table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="16%" class="tb_subtit"> 
                    <div align="center">UNIDADE 
                    </div></td>
                  <td width="16%" class="tb_subtit"> 
                    <div align="center">CURSO 
                    </div></td>
                  <td width="16%" class="tb_subtit"> 
                    <div align="center">ETAPA 
                    </div></td>
                  <td width="16%" class="tb_subtit"> 
                  	<div align="center">TURMA 
                  		</div></td>
                  <td width="16%" class="tb_subtit"><div align="center">DISCIPLINA</div></td>
                  <td width="16%" class="tb_subtit"> 
                  	<div align="center">PER&Iacute;ODO</div></td>
                  </tr>
                <tr> 
                  <td width="16%"> 
                    <div align="center"> 
                      <select name="unidade" id="unidade" class="select_style" onChange="recuperarCurso(this.value)">
                        <option value="999990"></option>
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
else
%>
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
                  </div></td>
                  <td width="16%"> 
                    <div align="center"> 
                      <div id="divCurso"> 
                        <select name="curso" class="select_style" id="curso">
                        <option value="999990" selected> 
                        </option>                        
                        </select>
                      </div>
                  </div></td>
                  <td width="16%"> 
                    <div align="center"> 
                      <div id="divEtapa"> 
                        <select name="etapa" class="select_style" id="etapa">
                        <option value="999990" selected> 
                        </option>                        
                        </select>
                      </div>
                  </div></td>
                  <td width="16%"> 
                  	<div align="center"> 
                  		<div id="divTurma"> 
                  			<select name="turma" class="select_style" id="turma">
                  				<option value="999990" selected> 
                  					</option>                        
                  				</select>
                  			</div>
                  		</div></td>
                  <td width="16%"><div align="center">
                  	<div id="divDisciplina">
                  		<select name="mat_prin" class="select_style" id="periodo">
                  			<option value="999999" selected> </option>
                  			</select>
                  		</div>
                  	</div></td>
                  <td width="16%"> 
                  	<div align="center"> 
                  		<div id="divPeriodo"> 
                  			<select name="periodo" class="select_style" id="periodo">
								<option value="0" selected></option>
                            								<%
							Set RS4 = Server.CreateObject("ADODB.Recordset")
							SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
							RS4.Open SQL4, CON0
					
							while not RS4.EOF
								NU_Periodo =  RS4("NU_Periodo")
								NO_Periodo= RS4("NO_Periodo")
								%>
								<option value="<%=NU_Periodo%>" >
									<%response.Write(NO_Periodo)%>
									</option>
								<%RS4.MOVENEXT
							WEND%>
								</select>
                  			</div>
                  		</div></td>
                  </tr>
                <tr>
                	<td height="15" colspan="6" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                		<tr>
                			<td class="tb_subtit"><div align="center">&nbsp;</div></td>
                			</tr>
                		<tr>
                			<td><div align="center"> <font class="form_dado_texto">De
                				<select name="dia_de" class="select_style">
                					<% for d=1 to 31
									
									d=d*1
                					if d = 1 then
										dia_selected="selected"
									else
										dia_selected=""									
									end if
									if d<10 then
										dia_exibe="0"&d
									else
										dia_exibe=d										
									end if %>									
                					<option value="<%response.Write(d)%>" <%response.Write(dia_selected)%>><%response.Write(dia_exibe)%></option>
                					
                					<%next%>	
                					</select>	/		
                				<select name="mes_de" id="mes_de" class="select_style">
                					<option value="1" selected>janeiro</option>
                					<option value="2">fevereiro</option>								
                					<option value="3">mar&ccedil;o</option>								
                					<option value="4">abril</option>								
                					<option value="5">maio</option>								
                					<option value="6">junho</option>								
                					<option value="7">julho</option>								
                					<option value="8">agosto</option>								
                					<option value="9">setembro</option>								
                					<option value="10">outubro</option>								
                					<option value="11">novembro</option>								
                					<option value="12">dezembro</option>								
                					<%end if%>
                					</select>/<select name="ano_de" class="select_style" id="ano_de">
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
                						</select>
                				
                				at&eacute; 
                					<select name="dia_ate" id="dia_ate" class="select_style">
                						<% for d=1 to 31
									dia_compara=dia_compara*1
									d=d*1

                					if dia_compara = d then
										dia_selected="selected"
									else
										dia_selected=""									
									end if
									if d<10 then
										dia_exibe="0"&d
									else
										dia_exibe=d										
									end if %>
                						<option value="<%response.Write(d)%>" <%response.Write(dia_selected)%>>
                							<%response.Write(dia_exibe)%>
                							</option>
                						<%next%>
                						</select>
                					/
	<select name="mes_ate" id="mes_ate" class="select_style">
		<%mes_compara=mes_compara*1
								if mes_compara="1" or mes_compara=1 then%>
		<option value="1" selected>janeiro</option>
		<% else%>
		<option value="1">janeiro</option>
		<%end if
								if mes_compara="2" or mes_compara=2 then%>
		<option value="2" selected>fevereiro</option>
		<% else%>
		<option value="2">fevereiro</option>
		<%end if
								if mes_compara="3" or mes_compara=3 then%>
		<option value="3" selected>mar&ccedil;o</option>
		<% else%>
		<option value="3">mar&ccedil;o</option>
		<%end if
								if mes_compara="4" or mes_compara=4 then%>
		<option value="4" selected>abril</option>
		<% else%>
		<option value="4">abril</option>
		<%end if
								if mes_compara="5" or mes_compara=5 then%>
		<option value="5" selected>maio</option>
		<% else%>
		<option value="5">maio</option>
		<%end if
								if mes_compara="6" or mes_compara=6 then%>
		<option value="6" selected>junho</option>
		<% else%>
		<option value="6">junho</option>
		<%end if
								if mes_compara="7" or mes_compara=7 then%>
		<option value="7" selected>julho</option>
		<% else%>
		<option value="7">julho</option>
		<%end if%>
		<%if mes_compara="8" or mes_compara=8 then%>
		<option value="8" selected>agosto</option>
		<% else%>
		<option value="8">agosto</option>
		<%end if
								if mes_compara="9" or mes_compara=9 then%>
		<option value="9" selected>setembro</option>
		<% else%>
		<option value="9">setembro</option>
		<%end if
								if mes_compara="10" or mes_compara=10 then%>
		<option value="10" selected>outubro</option>
		<% else%>
		<option value="10">outubro</option>
		<%end if
								if mes_compara="11" or mes_compara=11 then%>
		<option value="11" selected>novembro</option>
		<% else%>
		<option value="11">novembro</option>
		<%end if
								if mes_compara="12" or mes_compara=12 then%>
		<option value="12" selected>dezembro</option>
		<% else%>
		<option value="12">dezembro</option>
		<%end if%>
	</select>
                					/
	<select name="ano_ate" class="select_style" id="ano_ate">
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
	</select>
                					</font></div></td>
                			</tr>
                		</table></td>
                	</tr>
                <tr> 
                  <td height="15" colspan="6" bgcolor="#FFFFFF"><hr></td>
                </tr>
                <tr>
                  <td height="15" colspan="6" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0">
                    <tr>
                      <td width="33%"><div align="center"></div></td>
                      <td width="34%"><div align="center"></div></td>
                      <td width="33%"><div align="center"><font size="3" face="Courier New, Courier, mono">
                        <input type="submit" name="Submit2" value="Prosseguir" class="botao_prosseguir">
                      </font></div></td>
                    </tr>
                  </table></td>
                </tr>
            </table></td>
                            	</tr>
                          </table>
                    </form></td>
                </tr>
              </table>
		        </td>
        </tr>
      </table>
      </div> 
            </td>
          </tr>
        </table>        
      </form></td>
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