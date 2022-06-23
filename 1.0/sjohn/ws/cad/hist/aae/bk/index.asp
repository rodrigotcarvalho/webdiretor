<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<% 
session("nvg")=""
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
opt = request.QueryString("opt")
ano_info=nivel&"-"&nvg&"-"&ano_letivo
nvg=request.QueryString("nvg")
session("nvg")=nvg

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON7 = Server.CreateObject("ADODB.Connection") 
		ABRIR7 = "DBQ="& CAMINHO_h & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON7.Open ABRIR7		
		
		
		
call VerificaAcesso (CON,nvg,nivel)
autoriza=Session("autoriza")

call navegacao (CON,nvg,nivel)
navega=Session("caminho")

if opt="" or isnull("opt") then
	display="select"
elseif opt="search" then
	busca1=request.form("busca1") 
	busca2=request.form("busca2")
	ano_historico=request.form("ano_historico") 
	tipo_curso=request.form("tipo_curso")
	co_seg=request.form("co_seg") 	
	ordenacao=request.form("ordenacao")	
	display="redirect"	
	if busca1 <>"" or busca2 <>"" then		
		if busca1 ="" and busca2 <>"" then
			query = busca2
			busca_aluno="S"
		elseif busca1 <>"" then
			query = busca1 
			busca_aluno="S"
		else
			busca_aluno="N"						
		end if 
		
		if busca_aluno="S" then
			if IsNumeric(query) Then
			
				Set RSA = Server.CreateObject("ADODB.Recordset")
				SQLA = "SELECT CO_Matricula FROM TB_Alunos WHERE CO_Matricula = "& query
				RSA.Open SQLA, CON1	
				
				if RSA.EOF then
					display="reselect"
					dados_msg=query
					mensagem = 303
				else							
		  
					Set RS = Server.CreateObject("ADODB.Recordset")
					SQL = "SELECT * FROM TB_Historico_Ano where CO_Matricula = "& query
					RS.Open SQL, CON7
					
					if RS.EOF then
						incluir="S"
					else
						incluir="N"				
					end if
				end if	
				cod_cons = query
			ELSE
				Set RSA = Server.CreateObject("ADODB.Recordset")
				SQLA = "SELECT CO_Matricula FROM TB_Alunos WHERE NO_Aluno like '%"& query&"%'"
				RSA.Open SQLA, CON1
				
				if RSA.EOF then
					display="reselect"
					dados_msg=query
					mensagem = 304					
				else	
					alunos_encontrados=0
					while not RSA.EOF
						cod_cons = RSA("CO_Matricula")
						Set RS = Server.CreateObject("ADODB.Recordset")
						SQL = "SELECT * FROM TB_Historico_Ano where CO_Matricula = "& cod_cons
						RS.Open SQL, CON7
						
						if RS.EOF then
							incluir="S"
						else
							incluir="N"				
						end if
						if alunos_encontrados=0 then
							vetor_alunos=cod_cons
						else
							vetor_alunos=vetor_alunos&", "&cod_cons						
						end if						
						alunos_encontrados=alunos_encontrados+1						
					RSA.MOVENEXT
					WEND	
					
					if alunos_encontrados=1 then
						display="redirect"
					else
						display="list"
					end if
				end if	
			END IF		
		end if	
		obr=cod_cons
		if display="redirect" then		
			response.Redirect("resumo.asp?obr="&cod_cons&"&incl="&incluir)	
		end if			
	else
		obr=ano_historico&"$!$"&tipo_curso&"$!$"&co_seg&"$!$"&ordenacao
		response.Redirect("resumo.asp?obr="&obr&"&incl=N")				
	end if	
end if
%>
<html>
<head>
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
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
var currentlyActiveInputRef = false;
var currentlyActiveInputClassName = false;

function highlightActiveInput() {
  if(currentlyActiveInputRef) {
    currentlyActiveInputRef.className = currentlyActiveInputClassName;
  }
  currentlyActiveInputClassName = this.className;
  this.className = 'inputHighlighted';
  currentlyActiveInputRef = this;
}

function blurActiveInput() {
  this.className = currentlyActiveInputClassName;
}

function initInputHighlightScript() {
  var tags = ['INPUT','TEXTAREA'];
  for(tagCounter=0;tagCounter<tags.length;tagCounter++){
    var inputs = document.getElementsByTagName(tags[tagCounter]);
    for(var no=0;no<inputs.length;no++){
      if(inputs[no].className && inputs[no].className=='doNotHighlightThisInput')continue;
      if(inputs[no].tagName.toLowerCase()=='textarea' || (inputs[no].tagName.toLowerCase()=='input' && inputs[no].type.toLowerCase()=='text')){
        inputs[no].onfocus = highlightActiveInput;
        inputs[no].onblur = blurActiveInput;
      }
    }
  }
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
<% 
if display<>"list" then%>
  if (document.busca.busca1.value != "" && document.busca.busca2.value != "")
  {    alert("Por favor selecione SOMENTE uma opção de busca!")
    document.busca.busca1.value = "";
	document.busca.busca2.value = "";    
    document.busca.busca1.focus()
    return false
  }
  if (document.busca.busca1.value != "" && (document.busca.ano_historico.value != "999990" || document.busca.tipo_curso.value != "nulo"))
  {    alert("Por favor selecione SOMENTE uma opção de busca!")
    document.busca.busca1.value = "";
	document.busca.busca2.value = "";   
	var combo = document.getElementById("unidade");
	combo.options[0].selected = "true";
	//document.busca.unidade.selectedIndex = "999990";  
    document.busca.busca1.focus()
    return false
  }
  if (document.busca.busca2.value != "" && (document.busca.ano_historico.value != "999990" || document.busca.tipo_curso.value != "nulo"))
  {    alert("Por favor digite SOMENTE uma opção de busca!")
    document.busca.busca1.value = "";
	document.busca.busca2.value = "";   
	var combo = document.getElementById("unidade");
	combo.options[0].selected = "true";
    document.busca.busca1.focus()
    return false
  }  
	<%end if%>
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
   var f=document.forms[4]; 
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
						
						
						 function recuperarSegmento(tTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=s", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                    var resultado_t  = oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divSegmento.innerHTML =resultado_t

                                                           }
                                               }

                                               oHTTPRequest.send("t_pub=" + tTipo);
                                   }



 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
                        </script>
</head>
<% 
if display<>"list" then
	onload="onLoad=MM_callJS('document.busca.busca1.focus()')"
end if%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" <%response.Write(onload)%>>
         <form action="index.asp?opt=search&nvg=<%=nvg%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()">     
<%call cabecalho(nivel)
%>
<table width="1000" height="670" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr>             
    <td width="1000" height="10" colspan="5" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
    </td>
          </tr>
<%
if display="reselect" then%>
            <tr>              
    <td height="10" colspan="5"> 
      <%call mensagens(nivel,mensagem,1,dados_msg) %>
    </td>
			   </tr>          
<%
end if
if display="select" or display="reselect" then%>
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,300,0,0) %>
    </td>
			  </tr>	         	  
          <tr class="tb_tit">             
      <td height="10" colspan="5">Preencha um dos campos abaixo</td>
          </tr>
          <TR>
      <td height="10" valign="top"> 

                <table width="1000" border="0" cellpadding="0" cellspacing="0">
          <tr>          
            <td width="150"  height="10"> 
              <div align="right"><font class="form_dado_texto"> 
                Matr&iacute;cula:</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
            </strong></font></div></td>
            
            <td width="50" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="busca1" type="text" class="textInput" id="busca1" size="12">
              </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font></td>
            
            <td width="150" height="10"> 
              <div align="right"><font class="form_dado_texto"> Nome: </font></div></td>
            
            <td width="500" height="10" ><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="busca2" type="text" class="textInput" id="busca2" size="55" maxlength="50">
              </font></td>
            
            <td width="150" height="10">&nbsp;</td>
          </tr>
		  </table>
		  </td>
		  </TR>
           <tr>                   
    <td height="10" colspan="5" valign="top">  
<table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="250" align="center" class="tb_subtit"> 
                    ANO LETIVO
                  </td>
                  <td width="250" align="center" class="tb_subtit"> 
                    TIPO CURSO 
                    </td>
                  <td width="250" align="center" class="tb_subtit"> 
                    SEGMENTO</td>
                  <td width="250" align="center" class="tb_subtit"> 
                    ORDENA&Ccedil;&Atilde;O</td>
                </tr>
                <tr> 
                  <td width="250" align="center" > 
                      <select name="ano_historico" class="select_style" id="ano_historico">
                        <option value="999990" selected></option>
                        <%
		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT distinct DA_Ano FROM TB_Historico_Ano order by DA_Ano desc"
		RS7.Open SQL7, CON7
While not RS7.EOF
ano_historico = RS7("DA_Ano")
%>
                        <option value="<%response.Write(ano_historico)%>"> 
                        <%response.Write(ano_historico)%>
                        </option>
                        <%RS7.MOVENEXT
WEND
%>
                      </select>
                    </td>
                  <td width="250" align="center" > 
                        <select name="tipo_curso" class="select_style" id="tipo_curso" onChange="recuperarSegmento(this.value)">
                        <option value="nulo" selected></option>                        
                        <%		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Tipo_Curso order by NU_Ordem"
		RS0.Open SQL0, CON7
While not RS0.EOF
tipo_curso = RS0("TC_Curso")
no_abrv_curso = RS0("NO_Curso")
%>
                        <option value="<%response.Write(tipo_curso)%>"> 
                        <%response.Write(no_abrv_curso)%>
                        </option>
                        <%RS0.MOVENEXT
WEND
%>                        
                        </select></td>
                  <td width="250" align="center"> 
                     <div id="divSegmento"> 
                        <select name="co_seg"  class="select_style" id="co_seg">
                        </select>
                    </div></td>
                  <td width="250" align="center"> 
                        <select name="ordenacao" class="select_style" id="ordenacao" >
                        <option value="al" selected="selected"> 
                        Ano Letivo
                        </option>   
                        <option value="mt"> 
                        Matr&iacute;cula
                        </option>  
                        <option value="na"> 
                        Nome do Aluno
                        </option>    
                        <option value="es"> 
                        Escola
                        </option>                                                                                             
                    </select></td>
                </tr>
                <tr>
                  <td height="15" colspan="4" bgcolor="#FFFFFF"><hr></td>
                </tr>
                <tr> 
                  <td width="250" height="15" bgcolor="#FFFFFF"></td>
                  <td width="250" height="15" bgcolor="#FFFFFF"></td>
                  <td width="250" height="15" bgcolor="#FFFFFF"></td>
                  <td width="250" height="15" bgcolor="#FFFFFF"><font size="2" face="Arial, Helvetica, sans-serif">
                    <div align="center">
                      <input name="Submit" type="submit" class="botao_prosseguir" id="Submit" value="Procurar">
                    </div>
                  </font></td>
                </tr>
              </table>

    </td>
  </tr>
                <tr>                   
    <td colspan="5" valign="top"> 
    </td>
  </tr>        
<%
elseif display="list" then
	%>
                <tr>                   
    <td height="10" colspan="5" valign="top"> 
    <hr>
    </td>
  </tr>       
					<tr class="tb_corpo">                   
		<td height="10" colspan="5" class="tb_tit">Alunos Encontrados</td>
					</tr>
					<tr> 
					  
		<td colspan="5" valign="top">
         <ul> 
       <%	
		Set RSL = Server.CreateObject("ADODB.Recordset")
		SQLL = "SELECT * FROM TB_Alunos WHERE CO_Matricula in ("&vetor_alunos&") order by NO_Aluno"
		RSL.Open SQLL, CON1	
		   
		while not RSL.EOF
			cod_cons=RSL("CO_Matricula")
			nome = RSL("NO_Aluno")		
		
			Set RS = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT * FROM TB_Historico_Ano where CO_Matricula = "& cod_cons
			RS.Open SQL, CON7
			
			if RS.EOF then
				incluir="S"
			else
				incluir="N"				
			end if
			Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=ativos href=resumo.asp?obr="&cod_cons&"&incl="&incluir&">"&nome&"</a></font></li>")
		RSL.MOVENEXT
		WEND			
	%>
		  </ul>
          </td>
                </tr>  
<%
	END IF
%>                
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
</form>
</body>
<script type="text/javascript">
<!--
  initInputHighlightScript();
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