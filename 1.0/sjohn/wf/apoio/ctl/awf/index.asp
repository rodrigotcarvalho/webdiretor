<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%
opt = request.QueryString("opt")

ano_letivo_wf = request.QueryString("ano")
co_usr = session("co_user")
nivel=4
		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1

    	Set CON_WF = Server.CreateObject("ADODB.Connection") 
		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_WF.Open ABRIR_WF


if opt="a" then
nvg = session("chave")
co_usr = session("co_user")
chave=nvg
session("chave")=chave

nivel=4
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)


controle=request.form("wf")		

'co_apr1=request.form("apr1")
'co_apr2=request.form("apr2")
'co_apr3=request.form("apr3")
'co_apr4=request.form("apr4")
'co_apr5=request.form("apr5")
'co_apr6=request.form("apr6")
'co_prova1=request.form("prova1")
'co_prova2=request.form("prova2")
'co_prova3=request.form("prova3")
'co_prova4=request.form("prova4")
'co_prova5=request.form("prova5")
'co_prova6=request.form("prova6")

'response.Write(controle)
'response.End()

'outro=controle&co_apr1&co_prova1&co_apr2&co_prova2&co_apr3&co_prova3&co_apr4&co_prova4
outro="Alterou CO_controle do Web fam�lia para "& controle
sql_atualiza= "UPDATE TB_Controle SET CO_controle ='"&controle&"'"
Set RS4 = CON_WF.Execute(sql_atualiza)

else
nvg = request.QueryString("nvg")
chave=nvg
session("chave")=chave
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)
end if	

ano_info=nivel&"-"&chave&"-"&ano_letivo_wf


 call navegacao (CON,chave,nivel)
navega=Session("caminho")	

 Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Autoriz_Usuario_Grupo Where CO_Usuario = "&co_usr
		RS2.Open SQL2, CON
		
if RS2.EOF then

else		
co_grupo=RS2("CO_Grupo")
End if
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="file:../../../../img/mm_menu.js"></script>

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
function submitforminterno()  
{
   var f=document.forms[3]; 
      f.submit(); 
	  
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
//-->
</script>
<script>
<!--

// A fun��o abaixo pega a vers�o mais nova do xmlhttp do IE e verifica se � Firefox. Funciona nos dois.
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
// Cria��o do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicita��o HTTP. O primeiro par�metro informa o m�todo post/get
// O segundo par�metro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicita��o s�ncrona, o par�metro deve ser false
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=c3", true);
// Para solicita��es utilizando o m�todo post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A fun��o abaixo � executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto j� completou a solicita��o
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto � gerado no arquivo executa.asp e colocado no div
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
document.all.divControle.innerHTML =""
//recuperarEtapa()
                                                           }
                                               }
// Abaixo � enviada a solicita��o. Note que a configura��o
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarControle(cTipo)
                                   {
// Cria��o do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicita��o HTTP. O primeiro par�metro informa o m�todo post/get
// O segundo par�metro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicita��o s�ncrona, o par�metro deve ser false
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=ctrl", true);
// Para solicita��es utilizando o m�todo post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A fun��o abaixo � executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto j� completou a solicita��o
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto � gerado no arquivo executa.asp e colocado no div
                                                                       var resultado_ctrl= oHTTPRequest.responseText;
resultado_ctrl = resultado_ctrl.replace(/\+/g," ")
resultado_ctrl = unescape(resultado_ctrl)
document.all.divControle.innerHTML = resultado_ctrl																	   
                                                           }
                                               }
// Abaixo � enviada a solicita��o. Note que a configura��o
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("c_pub=" + cTipo);

                                   }		
								   
						// function GravaControle(variavel,valor)
						
function ConfirmaLiberaWebFamilia(){
	
	if (document.getElementById("wf").value == "L"){
		txt = "desbloqueio";
	}else {
		txt = "bloqueio";													   
	}
	msg_confirm = "Confirma o "+txt+" do Web Fam�lia";	
									   
	if (!confirm(msg_confirm)) {
	  return false;
	}		
	return true;
}
						
function Controle(avaliacao, periodo, mensagem,valor,id_rollback)

		   {
					   
				if (valor.value == "L"){
					txt = "desbloqueio"
					comando ="D"
				}else {
					txt = "bloqueio"
					comando ="L"													   
				}
				msg_confirm = "Confirma o "+txt+" "+avaliacao+" do " +periodo+ " do "+mensagem									   
				if (confirm(msg_confirm)) {

					GravarControle(valor.id);
				} else {
				  id_rollback.checked = true;
				}									   

		   }	
								   
function GravarControle(variavel)
			   {
						   var oHTTPRequest = createXMLHTTP();
						   oHTTPRequest.open("post", "bd.asp", true);
						   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
						   oHTTPRequest.onreadystatechange=function() {
									   if (oHTTPRequest.readyState==4){
												   var resultado_ctrl= oHTTPRequest.responseText;
resultado_ctrl = resultado_ctrl.replace(/\+/g," ")
resultado_ctrl = unescape(resultado_ctrl)
//document.all.divControle.innerHTML = resultado_ctrl	

//alert(resultado_ctrl)
									   }
						   }
							
						   oHTTPRequest.send("var_pub= " + variavel);

			   }													   								   
//-->
</script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0" >
<% call cabecalho (nivel)
	  %>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
                    
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
	  </td>
	  </tr>
      <%
if opt = "a" then%>
  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(nivel,9705,2,0)
%>
    </td>
                  </tr>
                  <% 	end if 

%>                  <tr> 
                    
    <td height="10"> 
      <%	call mensagens(nivel,9704,0,0) 
	  
		Set RS_WF = Server.CreateObject("ADODB.Recordset")
		SQL_WF = "SELECT * FROM TB_Controle"
		RS_WF.Open SQL_WF, CON_WF
		
controle=RS_WF("CO_controle")					  
%>
</td></tr>
<tr>

            <td valign="top"> 
                
        <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td width="653" height="15" class="tb_tit">Status do Web Fam&iacute;lia
</td>
          </tr>
          <tr> 
            <td valign="top">
           <FORM METHOD="POST" ACTION="index.asp?opt=a" onSubmit="return ConfirmaLiberaWebFamilia()">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="33%">&nbsp;</td>
                <td width="34%" class="form_corpo_12"><div align="center">
				<% 
				if controle="L" then
					response.Write("<input name=""wf"" type=""hidden"" id=""wf"" value=""D"">Web Familia est� no ar!")
					label="Tirar do Ar"
					classe="botao_excluir"
					
				else
					response.Write("<input name=""wf"" type=""hidden"" id=""wf"" value=""L"">Web Familia est� Fora do ar!")
					label="Colocar no Ar"	
					classe="botao_prosseguir"									
				end if							
				  %></div></td>
                <td width="33%"><div align="center"><INPUT TYPE=SUBMIT VALUE="<%response.Write(label)%>" class="<%response.Write(classe)%>"></div></td>
                </tr>
            </table></form></td>
          </tr>
          <tr>
            <td valign="top"><hr></td>
          </tr>
          <tr>
            <td valign="top" class="tb_tit">Status do Aproveitamento 
            Escolar</td>
          </tr>
          <tr>
            <td valign="top">
                
                  <table width="100%" border="0" cellspacing="0" cellpadding="0"> 
                  <tr><td class="cell_border_right_red">&nbsp;</td>                  
<%                  

                  
                  
                                  Set RSP = Server.CreateObject("ADODB.Recordset")
                SQLP = "SELECT Distinct(NU_Periodo),NO_Periodo FROM TB_Periodo order by NU_Periodo"
                RSP.Open SQLP, CON0
                        
                conta_periodo=0
                while not RSP.EOF				
                    nome_periodo=RSP("NO_Periodo")
                %>
                  <td colspan = "2" class="tb_subtit cell_border_right_red"><div align="center"><%response.Write(nome_periodo)%></div></td>
                <%
                conta_periodo=conta_periodo+1
                RSP.MOVENEXT
                WEND
%>                  
				</tr>	
                <tr>
				<td class="cell_border_right_red">&nbsp;</td>                  
                <% for cp=1 to conta_periodo %>
                  <td width="71" class="tb_subtit"><div align="center">Testes 
                  </div></td>
                  <td width="72" class="tb_subtit cell_border_right_red"><div align="center">Provas 
                  </div></td>
                <%next%> 
				</tr>                               		  
<%
				                                 
		

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0
	
While not RS0.EOF

	NU_Unidade = RS0("NU_Unidade")
	NO_Unidade = RS0("NO_Unidade")


exibe_nome_unidade = "N"

	
	if exibe_nome_unidade = "S" then	
%>
    <tr>
        <td colspan="<%response.Write(2+(conta_periodo*2))%>" valign="top" class="tb_tit">Nome da Unidade: <%response.write(NO_Unidade)%></td>
    </tr>	
<%	
end if

	Set RS1 = Server.CreateObject("ADODB.Recordset")
	SQL1 = "SELECT CO_Curso, CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&NU_Unidade&" ORDER BY CO_Curso, CO_Etapa"
	RS1.Open SQL1, CON0		

	CO_Curso_CTRL = "NULO"
		
	While not RS1.EOF		
	
		CO_Curso = RS1("CO_Curso")
	
		if CO_Curso_CTRL <> CO_Curso then
			CO_Curso_CTRL = CO_Curso
			Set RS1a = Server.CreateObject("ADODB.Recordset")
			SQL1a = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
			RS1a.Open SQL1a, CON0
			
			if RS1a.EOF then
				NO_Curso = ""
				CO_Conc = ""
			else
				NO_Curso = RS1a("NO_Curso")	
				CO_Conc = RS1a("CO_Conc")	
			end if			
		end if	
    
		co_etapa = RS1("CO_Etapa")

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&CO_Curso&"' AND CO_Etapa='"&co_etapa&"'"
		RS0c.Open SQL0c, CON0
		
		no_etapa = RS0c("NO_Etapa")		
	
	
		Set RS_WF = Server.CreateObject("ADODB.Recordset")
		SQL_WF = "SELECT * FROM TB_Autoriza_WF WHERE NU_Unidade="&NU_Unidade&" AND CO_Curso='"&CO_Curso&"' and CO_Etapa='"&co_etapa&"'"
		RS_WF.Open SQL_WF, CON_WF
		
		if not RS_WF.EOF then
					
				co_apr1=RS_WF("CO_apr1")
				co_apr2=RS_WF("CO_apr2")
				co_apr3=RS_WF("CO_apr3")
				co_apr4=RS_WF("CO_apr4")
				co_apr5=RS_WF("CO_apr5")
				co_apr6=RS_WF("CO_apr6")
				co_apr7=RS_WF("CO_apr7")					
				co_prova1=RS_WF("CO_prova1")
				co_prova2=RS_WF("CO_prova2")
				co_prova3=RS_WF("CO_prova3")
				co_prova4=RS_WF("CO_prova4")
				co_prova5=RS_WF("CO_prova5")		
				co_prova6=RS_WF("CO_prova6")	
				co_prova7=RS_WF("CO_prova7")
			
				
			
		%>
		  <tr>
                <td class="tb_subtit cell_border_right_red" width="100"><div align="center">                    
                    <%response.Write(NO_Curso&"<br>"&no_etapa)%></div>
                </td>          
          <%
		  for cp=1 to conta_periodo 	
		  
		  
		Set RS0p = Server.CreateObject("ADODB.Recordset")
		SQL0p = "SELECT * FROM TB_Periodo where NU_Periodo="&cp
		RS0p.Open SQL0p, CON0
		
		nome_periodo = RS0p("NO_Periodo")	
                                    
                    if cp=1 then				
                        if co_apr1="L" then
                            apr_checked_lib ="checked"
                            apr_checked_blq =""														
                        else
                            apr_checked_lib =""
                            apr_checked_blq ="checked"																	
                        end if
                        if co_prova1="L" then
                            pr_checked_lib ="checked"
                            pr_checked_blq =""														
                        else
                            pr_checked_lib =""
                            pr_checked_blq ="checked"																	
                        end if			
                                        
                    elseif cp=2 then
                        if co_apr2="L" then
                            apr_checked_lib	="checked"
                            apr_checked_blq	=""														
                        else
                            apr_checked_lib	=""
                            apr_checked_blq	="checked"																	
                        end if
                        if co_prova2="L" then
                            pr_checked_lib	="checked"
                            pr_checked_blq	=""														
                        else
                            pr_checked_lib	=""
                            pr_checked_blq	="checked"																	
                        end if												
                    elseif cp=3 then
                        if co_apr3="L" then
                            apr_checked_lib	="checked"
                            apr_checked_blq	=""														
                        else
                            apr_checked_lib	=""
                            apr_checked_blq	="checked"																	
                        end if
                        if co_prova3="L" then
                            pr_checked_lib	="checked"
                            pr_checked_blq	=""														
                        else
                            pr_checked_lib	=""
                            pr_checked_blq	="checked"																	
                        end if	
                
                    elseif cp=4 then
                        if co_apr4="L" then
                            apr_checked_lib	="checked"
                            apr_checked_blq	=""														
                        else
                            apr_checked_lib	=""
                            apr_checked_blq	="checked"																	
                        end if
                        if co_prova4="L" then
                            pr_checked_lib	="checked"
                            pr_checked_blq	=""														
                        else
                            pr_checked_lib	=""
                            pr_checked_blq	="checked"																	
                        end if					
            
                    elseif cp=5 then
                        if co_apr5="L" then
                            apr_checked_lib	="checked"
                            apr_checked_blq	=""														
                        else
                            apr_checked_lib	=""
                            apr_checked_blq	="checked"																	
                        end if
                        if co_prova5="L" then
                            pr_checked_lib	="checked"
                            pr_checked_blq	=""														
                        else
                            pr_checked_lib	=""
                            pr_checked_blq	="checked"																	
                        end if						
                    elseif cp=6 then
                        if co_apr6="L" then
                            apr_checked_lib	="checked"
                            apr_checked_blq	=""														
                        else
                            apr_checked_lib	=""
                            apr_checked_blq	="checked"																	
                        end if
                        if co_prova6="L" then
                            pr_checked_lib	="checked"
                            pr_checked_blq	=""														
                        else
                            pr_checked_lib	=""
                            pr_checked_blq	="checked"																	
                        end if		
                    elseif cp=7 then
                        if co_apr7="L" then
                            apr_checked_lib	="checked"
                            apr_checked_blq	=""														
                        else
                            apr_checked_lib	=""
                            apr_checked_blq	="checked"																	
                        end if
                        if co_prova7="L" then
                            pr_checked_lib	="checked"
                            pr_checked_blq	=""														
                        else
                            pr_checked_lib	=""
                            pr_checked_blq	="checked"																	
                        end if	
                    end if	
					%>            	
                  <td width="71"><div align="center">
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                       <tr>
                        <td>
<input type="radio" name="<%response.Write("NOME_"&NU_Unidade&"_"&CO_Curso&"_"&co_etapa&"_t_"&cp)%>" id="<%response.Write("ID_"&NU_Unidade&"_"&CO_Curso&"_"&co_etapa&"_t_b_"&cp)%>" value="D" class="borda" <%response.Write(apr_checked_blq)%> onclick="Controle('do Teste', '<%response.write(nome_periodo)%>', '<%response.write(no_etapa&" "&CO_Conc&" "&NO_Curso)%>', this, <%response.Write("ID_"&NU_Unidade&"_"&CO_Curso&"_"&co_etapa&"_t_l_"&cp)%>)"/></td>                        
                        <td class="form_dado_texto">Bloq</td>
                      </tr>
                      <tr>
                        <td>
                        <input type="radio" name="<%response.Write("NOME_"&NU_Unidade&"_"&CO_Curso&"_"&co_etapa&"_t_"&cp)%>" id="<%response.Write("ID_"&NU_Unidade&"_"&CO_Curso&"_"&co_etapa&"_t_l_"&cp)%>" value="L" class="borda" <%response.Write(apr_checked_lib)%> onclick="Controle('do Teste', '<%response.write(nome_periodo)%>', '<%response.write(no_etapa&" "&CO_Conc&" "&NO_Curso)%>', this, <%response.Write("ID_"&NU_Unidade&"_"&CO_Curso&"_"&co_etapa&"_t_b_"&cp)%>)"/>
                        </td>
                        
                        <td class="form_dado_texto">Lib</td>
                      </tr>
                    </table>
                  </div></td>
                  <td width="72" class="cell_border_right_red"><div align="center">
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><input type="radio" name="<%response.Write("NOME_"&NU_Unidade&"_"&CO_Curso&"_"&co_etapa&"_p_"&cp)%>" id="<%response.Write("ID_"&NU_Unidade&"_"&CO_Curso&"_"&co_etapa&"_p_b_"&cp)%>" value="D" class="borda" <%response.Write(pr_checked_blq)%> onclick="Controle('da Prova', '<%response.write(nome_periodo)%>', '<%response.write(no_etapa&" "&CO_Conc&" "&NO_Curso)%>', this, <%response.Write("ID_"&NU_Unidade&"_"&CO_Curso&"_"&co_etapa&"_p_l_"&cp)%>)"/></td>
                        <td class="form_dado_texto">Bloq</td>
                      </tr>
                      <tr>
                        <td><input type="radio" name="<%response.Write("NOME_"&NU_Unidade&"_"&CO_Curso&"_"&co_etapa&"_p_"&cp)%>" id="<%response.Write("ID_"&NU_Unidade&"_"&CO_Curso&"_"&co_etapa&"_p_l_"&cp)%>" value="L" class="borda" <%response.Write(pr_checked_lib)%> onclick="Controle('da Prova', '<%response.write(nome_periodo)%>', '<%response.write(no_etapa&" "&CO_Conc&" "&NO_Curso)%>', this, <%response.Write("ID_"&NU_Unidade&"_"&CO_Curso&"_"&co_etapa&"_p_b_"&cp)%>)"/></td>
                        <td class="form_dado_texto">Lib</td>
                      </tr>
                    </table>
                  </div></td>                   
                  <%next%>

                </tr>
                <tr><td colspan="<%response.Write(1+(conta_periodo*2))%>"><hr/></td></tr>             

    <%	end if
	RS1.MOVENEXT
    WEND
 RS0.MOVENEXT
WEND
%>                
            

              <tr>
                <td><div id="divControle">
                  
                </div></td>
              </tr>
            </table></td>
          </tr>
        </table>
              </td>
  </tr>
		  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

</body>
</html>
<%
if opt="a" then
			call GravaLog (chave,outro)
end if
If Err.number<>0 then
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