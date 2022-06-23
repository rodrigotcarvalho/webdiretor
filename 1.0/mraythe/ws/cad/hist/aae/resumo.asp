<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<% 
Session.LCID = 1046
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

Set CON = Server.CreateObject("ADODB.Connection") 
ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
CON.Open ABRIR

Set CON1 = Server.CreateObject("ADODB.Connection") 
ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
CON1.Open ABRIR1

Set CON2 = Server.CreateObject("ADODB.Connection") 
ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
CON2.Open ABRIR2

Set CON7 = Server.CreateObject("ADODB.Connection") 
ABRIR7 = "DBQ="& CAMINHO_h & ";Driver={Microsoft Access Driver (*.mdb)}"
CON7.Open ABRIR7		
		
Set CON0 = Server.CreateObject("ADODB.Connection") 
ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
CON0.Open ABRIR0	
		
call navegacao (CON,chave,nivel)
navega=Session("caminho")	
voltar= request.QueryString("voltar")	
opt= request.QueryString("opt")	
if voltar="S" then
	obr= Session("obr_AAC")
	incl= Session("incl_AAC")
else		
	obr= request.QueryString("obr")
	incl= request.QueryString("incl")
end if	
Session("obr_AAC")=obr
Session("incl_AAC")=incl
dados_obr=split(obr,"$!$")	
if ubound(dados_obr)=0 then
	if incl="S" and voltar<>"S" then
		RESPONSE.Redirect("incluir.asp?cod="&obr&"&opt=inc")
	end if
	cod_aluno= dados_obr(0)
	
	Set RSA = Server.CreateObject("ADODB.Recordset")
	SQLA = "SELECT NO_Aluno FROM TB_Alunos WHERE CO_Matricula = "& cod_aluno
	RSA.Open SQLA, CON1	
	
	if RSA.EOF then
		nome_aluno = "Nome não cadastrado"
	else
		nome_aluno = RSA("NO_Aluno")
	end if	
	
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Historico_Ano where CO_Matricula = "& cod_aluno
	RS.Open SQL, CON7	
	
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Matriculas WHERE CO_Matricula ="& cod_aluno&" and NU_Ano in (SELECT MAX(NU_Ano) FROM TB_Matriculas WHERE CO_Matricula ="& cod_aluno&")"
	RS.Open SQL, CON1

	if RS.EOF then
	
	else
		ano_aluno = RS("NU_Ano")
		rematricula = RS("DA_Rematricula")
		situacao = RS("CO_Situacao")
		encerramento= RS("DA_Encerramento")
		unidade= RS("NU_Unidade")
		curso= RS("CO_Curso")
		etapa= RS("CO_Etapa")
		turma= RS("CO_Turma")
		cham= RS("NU_Chamada")
			
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
		RSCONTST.Open SQLCONTST, CON0
							
		no_situacao = RSCONTST("TX_Descricao_Situacao")	
		
		call GeraNomes("PORT",unidade,curso,etapa,CON0)
		no_unidades = session("no_unidades")
		no_grau = session("no_grau")
		no_serie = session("no_serie")		
	end if	
	
	tipo_consulta = "A"	
	ordena=" ORDER BY DA_Ano desc"
else
	ano_historico= dados_obr(0)
	tipo_curso= dados_obr(1)
	co_seg= dados_obr(2)
	ordem= dados_obr(3)
	
	Select case ordem
	
	case "al"
	ordena=" ORDER BY DA_Ano desc"
	
	case "mt"
	ordena=" ORDER BY CO_Matricula"
	
	case "es"
	ordena=" ORDER BY NO_Escola"
	
	case "na"
	ordena=""
	
	end select	
	tipo_consulta = "UCET"	
end if
Session("ordena_AEE") = ordem
%>
<html>
<head>
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../../../js/mm_menu.js"></script>
<script type="text/javascript" src="../../../../js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
<!--
var checkflag = "false";
function check(field) {
if (checkflag == "false") {
for (i = 0; i < field.length; i++) {
field[i].checked = true;}
checkflag = "true";
return "Desmarcar Todos"; }
else {
for (i = 0; i < field.length; i++) {
field[i].checked = false; }
checkflag = "false";
return "Marcar Todos"; }
}

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
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function checksubmit()
{
<% if tipo_consulta = "A"	then%>	
  if (document.busca.busca1.value != "" && document.busca.busca2.value != "")
  {    alert("Por favor digite SOMENTE uma opção de busca!")
    document.busca.busca1.value = "";
	document.busca.busca2.value = "";    
    document.busca.busca1.focus()
    return false
  }
//  if (document.busca.busca1.value != "" && document.busca.ano_historico.value != "999990")
//  {    alert("Por favor digite SOMENTE uma opção de busca!")
//    document.busca.busca1.value = "";
//	document.busca.busca2.value = "";   
//	var combo = document.getElementById("unidade");
//	combo.options[0].selected = "true";
//	//document.busca.unidade.selectedIndex = "999990";  
//    document.busca.busca1.focus()
//    return false
//  
// }
  if (document.busca.ano_historico.value != "999990" && document.busca.busca2.value != "")
  {    alert("Por favor digite SOMENTE uma opção de busca!")
    document.busca.busca1.value = "";
	document.busca.busca2.value = "";   
	var combo = document.getElementById("unidade");
	combo.options[0].selected = "true";
    document.busca.busca1.focus()
    return false
  }  
<%end if%>  
//    if (document.busca.busca1.value == "" && document.busca.busca2.value == "" && document.busca.ano_historico.value == 999990)	
//  {
//	  alert("Por favor digite uma opção de busca!")
//		document.busca.busca1.focus()
//		return false
//  }
  return true
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

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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


function checkTheBox() {
   var chk = document.getElementsByName('historico')
    var len = chk.length

    for(i=0;i<len;i++)
    {
         if(chk[i].checked){
        return true;
          }
    }
	alert("Pelo menos um histórico deve ser selecionado!")		
    return false;
    }
	
 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}

function checkTheBoxRedireciona() {
	

      if (!checkTheBox()){	
		 return false;
	  } else {  
        document.forms[0].submit(); 
	  } 
}
                        </script>
</head>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"> 
<form action="redireciona.asp" method="post" name="busca" id="busca" onSubmit="return checkTheBox()">
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
   <% if opt="ok1" then    %>    
            <tr>    
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,418,2,0) %>
    </td>
			  </tr>
<% end if
 if opt="ok2" then    %>    
            <tr>    
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,419,2,0) %>
    </td>
			  </tr>
	<% end if
	 if opt="ok3" then    %>    
            <tr>    
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,420,2,0) %>
    </td>
			  </tr>			  			  
<% end if%>			  		  
	  
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,9707,0,0) %>
    </td>
			  </tr>			  
          <tr>      
    <td valign="top"> 
      <table width="1000" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo"
>
        <tr> 
          <td width="653" class="tb_tit"
>Crit&eacute;rios Informados</td>
          <td width="113" class="tb_tit"
> </td>
        </tr>

          <tr> 
            <td colspan="2" valign="top"> 
            <%if tipo_consulta = "UCET" then%>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="tb_subtit">
                <td width="250" height="10" align="center"> ANO LETIVO </td>
                <td width="250" align="center"> TIPO CURSO </td>
                <td width="250" align="center"> SEGMENTO</td>
                <td width="250" align="center"> ORDENA&Ccedil;&Atilde;O</td>
                </tr>
              <tr class="tb_corpo">
                <td width="250" height="10"><div align="center"> <font class="form_dado_texto">
				  <select name="ano_historico" class="select_style" id="ano_historico" disabled>
					<option value="999990"></option>                  
                        <%
		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT distinct DA_Ano FROM TB_Historico_Ano order by DA_Ano desc"
		RS7.Open SQL7, CON7
ano_historico = ano_historico*1		
While not RS7.EOF
ano_historico_bd = RS7("DA_Ano")
ano_historico_bd = ano_historico_bd*1
if ano_historico_bd = ano_historico then
	selected_ano = "selected"
else
	selected_ano=""		
end if	

%>
                        <option value="<%response.Write(ano_historico_bd)%>" <%response.Write(selected_ano)%>> 
                        <%response.Write(ano_historico_bd)%>
                        </option>
                        <%RS7.MOVENEXT
WEND
%>
                  </select></font></div></td>
                <td width="250" height="10" align="center">
<select name="tipo_curso" class="select_style" id="tipo_curso" onChange="recuperarSegmento(this.value)" disabled>  
<option value="nulo" selected></option>                  
                        <%		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Tipo_Curso order by NU_Ordem"
		RS0.Open SQL0, CON7
		
While not RS0.EOF
tipo_curso_bd = RS0("TC_Curso")
no_abrv_curso = RS0("NO_Curso")
if tipo_curso_bd = tipo_curso then
	selected_curso = "selected"
else
	selected_curso=""	
end if	
%>
                        <option value="<%response.Write(tipo_curso_bd)%>" <%response.Write(selected_curso)%>> 
                        <%response.Write(no_abrv_curso)%>
                        </option>
                        <%RS0.MOVENEXT
WEND
%>                        
                        </select></td>
                  <td width="250" align="center"> 
                     <div id="divSegmento"> 
                    <select name="co_seg" class="select_style" id="co_seg" disabled>
					<option value="nulo" selected></option>                      
                    <%                   
                        Set RS0 = Server.CreateObject("ADODB.Recordset")
                        SQL0 = "SELECT * FROM TB_Segmento where TP_Curso='"&tipo_curso&"' order by NU_Ordem"
                        RS0.Open SQL0, CON7
                    
                    
                        While not RS0.EOF
                        
                        co_seqmento = RS0("CO_Seg")		
                        no_seqmento = RS0("NO_Abreviado_Curso")	
						if co_seg = co_seqmento then
							selected_seg = "selected"
						else
							selected_seg=""	
						end if								
                        %>
                                                <option value="<%response.Write(co_seqmento)%>" <%response.Write(selected_seg)%>> 
                                                <%response.Write(no_seqmento)%>
                                                </option>
                                                <%
                        RS0.MOVENEXT
                        WEND
                    %>
                    </select>
                    </div></td>
                  <td width="250" align="center"> 
                        <select name="ordenacao" class="select_style" id="ordenacao" disabled>
<%

if ordem ="al" then
	selected_al = "selected"
elseif ordem ="mt" then
	selected_mt = "selected"	
elseif ordem ="na" then
	selected_na = "selected"
elseif ordem ="es" then
	selected_es = "selected"		
end if	

%>                        
                        <option value="al" <%response.Write(selected_al)%>> 
                        Ano Letivo
                        </option>   
                        <option value="mt" <%response.Write(selected_mt)%>> 
                        Matr&iacute;cula
                        </option>  
                        <option value="na" <%response.Write(selected_na)%>> 
                        Nome do Aluno
                        </option>    
                        <option value="es" <%response.Write(selected_es)%>> 
                        Escola
                        </option>                                                                                             
                    </select>
                </td>                
                </tr>
              <tr class="tb_corpo">
                <td height="10" colspan="4"><hr></td>
                </tr>
              <tr class="tb_corpo">
                <td height="10" colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="33%">&nbsp;</td>
                    <td width="34%">&nbsp;</td>
                    <td width="33%" align="center"><font size="2" face="Arial, Helvetica, sans-serif">
                      <input name="Button" type="button" class="botao_prosseguir" id="Submit" value="Procurar">
                    </font></td>
                  </tr>
                </table></td>
              </tr>
            </table>
<%else%>	
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td height="10"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="155" height="10"><div align="right"><font class="form_dado_texto"> Matr&iacute;cula: </font></div></td>
                      <td width="78" height="10"><font class="form_dado_texto">
                        </font><font size="2" face="Arial, Helvetica, sans-serif">
                        <input name="busca1" type="text" disabled class="textInput" id="busca1" value="<%response.Write(cod_aluno)%>" size="12">
                        </font></td>
                      <td width="49" height="10"><div align="right"><font class="form_dado_texto"> Nome: </font></div></td>
                      <td width="710" height="10">
                        <input name="busca2" type="text" disabled class="textInput" id="busca2" value="<%response.Write(nome_aluno)%>" size="55" maxlength="50">
                        </font></td>
                      </tr>
                    <tr>
                      <td height="10" colspan="4"><hr></td>
                      </tr>
                    <tr>
                      <td height="10" colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td width="33%">&nbsp;</td>
                          <td width="34%">&nbsp;</td>
                          <td width="33%" align="center"><font size="2" face="Arial, Helvetica, sans-serif">
                            <input name="Button" type="button" class="botao_prosseguir" id="Submit" value="Procurar">
                          </font></td>
                        </tr>
                      </table></td>
                    </tr>
                    <tr>
                      <td height="10" colspan="4">&nbsp;</td>
                    </tr>
                  </table></td>
                </tr>
                <!--
                <tr class="tb_tit"> 
                  <td height="10"> Dados Escolares</td>
                </tr>
                <tr>
                  <td height="10"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr class="tb_subtit">
                      <td height="10"><div align="center">
                        Ano</div></td>
                      <td height="10"><div align="center">Matr&iacute;cula</div></td>
                      <td height="10"><div align="center">Cancelamento</div></td>
                      <td height="10"><div align="center"> Situa&ccedil;&atilde;o</div></td>
                      <td height="10"><div align="center">Unidade</div></td>
                      <td height="10"><div align="center">Curso</div></td>
                      <td height="10"><div align="center"> Etapa</div></td>
                      <td height="10"><div align="center">Turma </div></td>
                      <td height="10"><div align="center">Chamada</div></td>
                      </tr>
                    <tr class="tb_corpo">
                      <td height="10"><div align="center"> <font class="form_dado_texto">
                        <%'response.Write(ano_aluno)%>
                        </font></div></td>
                      <td height="10"><div align="center"> <font class="form_dado_texto">
                        <input name="cod2" type="hidden" id="cod" value="<%'response.Write(codigo)%>">
                        <%'response.Write(rematricula)%>
                        </font></div></td>
                      <td height="10"><div align="center"> <font class="form_dado_texto">
                        <%'response.Write(encerramento)%>
                        </font></div></td>
                      <td height="10"><div align="center"> <font class="form_dado_texto">
                        <%					
					'response.Write(no_situacao)%>
                        </font></div></td>
                      <td height="10"><div align="center"> <font class="form_dado_texto">
                        <%'response.Write(no_unidades)%>
                        </font></div></td>
                      <td height="10"><div align="center"> <font class="form_dado_texto">
                        <%'response.Write(no_grau)%>
                        </font></div></td>
                      <td height="10"><div align="center"> <font class="form_dado_texto">
                        <%'response.Write(no_serie)%>
                        </font></div></td>
                      <td height="10"><div align="center"> <font class="form_dado_texto">
                        <%'response.Write(turma)%>
                        </font></div></td>
                      <td height="10"><div align="center"> <font class="form_dado_texto">
                        <%'response.Write(cham)%>
                        </font></div></td>
                      </tr>
                  </table></td>
                </tr>-->
              </table>
              <%end if%>
            </td>
        </tr>
        <tr height="10"> 
          <td height="10" colspan="2" >&nbsp;</td>
        </tr>
        <tr> 
          <td height="10" colspan="2" ></td>
        </tr>		
        <tr> 
          <td height="10" colspan="2" class="tb_tit"
>Anos Letivos Registrados</td>
        </tr>
        <tr > 
          <td height="154" colspan="2" valign="top"> 
	  
		  
		  
              <table width="1000" border="0" cellspacing="0" cellpadding="0">
                <tr class="tb_subtit"> 
                  <td width="20" height="10"> <input type="checkbox" name="todos" class="borda" value="" onClick="this.value=check(this.form.historico)"> 
                  </td>
                  <td width="50" align="center">Ano<br>Letivo</td>
                  <td width="30" align="center">Seq</td>
                  <td width="60" align="center">Matr&iacute;cula</td>
                  <td width="280" align="left">&nbsp;Nome</td>
                  <td width="40" align="center">Curso</td>
                  <td width="110" align="center">Etapa</td>
                  <td width="180" align="center">Escola</td>
                  <td width="80" align="center">Situa&ccedil;&atilde;o</td>
                  <td width="80" align="center">Tipo de Registro</td>
                  <td width="60" align="center">Alterado em</td>
                </tr>
				                <tr> 
                  <td colspan="11"><hr width="1000"></td>
                </tr>
                <%	
				registros_encontrados = "N"	
		Set RS = Server.CreateObject("ADODB.Recordset")					
		if tipo_consulta = "A"	then
			SQL = "SELECT * FROM TB_Historico_Ano where CO_Matricula = "& cod_aluno&ordena	
		else
			if ano_historico<>"999990" then
				sql_dinamico = " DA_Ano = "& ano_historico
			end if				
			if tipo_curso <> "nulo" then
				if isnull(sql_dinamico) or sql_dinamico="" then
				
				else
					sql_dinamico = sql_dinamico&" AND "
				end if				
				sql_dinamico =  sql_dinamico&"TP_Curso ='"&tipo_curso&"'"
			end if		
			if isnull(co_seg) or co_seg="" or co_seg = "nulo" then
			else
				if isnull(sql_dinamico) or sql_dinamico="" then
				
				else
					sql_dinamico = sql_dinamico&" AND "
				end if					
				sql_dinamico = sql_dinamico&" CO_Seg ='"&co_seg&"'"		
			end if
			if isnull(sql_dinamico) or sql_dinamico="" then
			
			else
				sql_dinamico = " where "&sql_dinamico
			end if	
			SQL = "SELECT * FROM TB_Historico_Ano "&sql_dinamico&ordena
		end if				
	RS.Open SQL, CON7

	Set Rs_ordena = Server.CreateObject ( "ADODB.RecordSet" )
	'200 -> VarChar (String), 7 -> Data, 139 -> Numeric
	Rs_ordena.Fields.Append "ano_historico", 139, 10
	Rs_ordena.Fields.Append "seq", 139, 10	
	Rs_ordena.Fields.Append "matric", 139, 10
	Rs_ordena.Fields.Append "nome", 200, 255		
	Rs_ordena.Fields.Append "curso", 200, 3
	Rs_ordena.Fields.Append "segmento", 200, 5
	Rs_ordena.Fields.Append "escola", 200, 50
	Rs_ordena.Fields.Append "situac", 200, 1
	Rs_ordena.Fields.Append "tipo_regist", 200, 1
	Rs_ordena.Fields.Append "data", 7
	Rs_ordena.Open
	if RS.EOF then
		registros_encontrados = "N"	
	%>
       <tr class="<%=cor%>">         
          <td width="20">&nbsp;</td>
          <td colspan="10" align="center" class="form_dado_texto">N&atilde;o existem hist&oacute;ricos cadastrados que atendam os par&acirc;metros informados</td>
        </tr>
	<%else
		registros_encontrados = "S"	
		tot_rec=0	
		 WHILE NOT RS.EOF
			ano_hist = RS("DA_Ano")		 
			seq_hist = RS("NU_Seq")
			matric_hist = RS("CO_Matricula") 
			curso_hist = RS("TP_Curso")
			seg_hist = RS("CO_Seg")
			escola_hist = RS("NO_Escola")
			situac_hist = RS("IN_Aprovado")
			tp_reg_hist = RS("TP_Registro")
			data_hist = RS("DT_Registro")			 		 
			 
			if data_hist = "" or isnull(data_hist) then
				data_hist = "31/12/9999"
			end if
			if seg_hist = "" or isnull(seg_hist) then			
				seg_hist = ""
			end if			
			 
			Set RSN = Server.CreateObject("ADODB.Recordset")
			SQLN = "SELECT * FROM TB_Alunos where CO_Matricula ="& matric_hist
			RSN.Open SQLN, CON1	
			if RSN.eof then
				nome_hist = "ZZZNão cadastrado"
			else
				nome_hist = RSN("NO_Aluno")
			end if		
			
			Rs_ordena.AddNew			
			Rs_ordena.Fields("ano_historico").Value = ano_hist			
			Rs_ordena.Fields("seq").Value = seq_hist
			Rs_ordena.Fields("matric").Value = matric_hist
			Rs_ordena.Fields("nome").Value = nome_hist
			Rs_ordena.Fields("curso").Value = curso_hist
			Rs_ordena.Fields("segmento").Value = seg_hist
			Rs_ordena.Fields("escola").Value = escola_hist
			Rs_ordena.Fields("situac").Value = situac_hist
			Rs_ordena.Fields("tipo_regist").Value = tp_reg_hist
			Rs_ordena.Fields("data").Value = data_hist	
			tot_rec=tot_rec+1		
		RS.MOVENEXT
		WEND								 		
			if ordem = "na" then
				Rs_ordena.Sort = "nome ASC"
			end if
			Rs_ordena.PageSize = 30
			 
			if Request.QueryString("pagina")="" then
				  intpagina = 1
				  Rs_ordena.MoveFirst
			else
				if cint(Request.QueryString("pagina"))<1 then
					intpagina = 1
				else
					if cint(Request.QueryString("pagina"))>Rs_ordena.PageCount then  
						intpagina = Rs_ordena.PageCount
					else
						intpagina = Request.QueryString("pagina")
					end if
				end if   
			 end if   
		
			Rs_ordena.AbsolutePage = intpagina
			intrec = 0
			check=2
			While intrec<Rs_ordena.PageSize and Not Rs_ordena.EoF
				if check mod 2 =0 then
					cor = "tb_fundo_linha_par" 
				else 
					cor ="tb_fundo_linha_impar"
				end if	
				 
				 
				ano_exibe = Rs_ordena.Fields("ano_historico").Value 			
				seq_exibe = Rs_ordena.Fields("seq").Value 
				matric_exibe = Rs_ordena.Fields("matric").Value 
				nome_exibe = Rs_ordena.Fields("nome").Value 
				curso_exibe = Rs_ordena.Fields("curso").Value 
				seg_bd = Rs_ordena.Fields("segmento").Value 
				escola_exibe = Rs_ordena.Fields("escola").Value 
				situac_exibe = Rs_ordena.Fields("situac").Value 
				tp_reg_exibe = Rs_ordena.Fields("tipo_regist").Value 
				data_exibe = Rs_ordena.Fields("data").Value 		
				
				optobr = ano_exibe&"$!$"&seq_exibe&"$!$"&matric_exibe						
				
				if isnull(curso_exibe) or isnull(seg_exibe) then
				
				else
					Set RS0 = Server.CreateObject("ADODB.Recordset")
					SQL0 = "SELECT * FROM TB_Segmento where TP_Curso='"&curso_exibe&"' AND CO_Seg='"&seg_bd&"' order by NU_Ordem"
					RS0.Open SQL0, CON7
					
					if RS0.EOF then
						seg_exibe = ""
					else
						seg_exibe = RS0("NO_Abreviado_Curso")
					end if		
				end if	
				
				if isnull(situac_exibe) or isnull(situac_exibe) then
				
				else
					Set RS0 = Server.CreateObject("ADODB.Recordset")
					SQL0 = "SELECT * FROM TB_Resultado_Final where TP_Resultado="&situac_exibe
					RS0.Open SQL0, CON7
					situac_exibe = RS0("NO_Resultado")	
				end if								
				
				if tp_reg_exibe = "M" then
					tp_reg_exibe = "Manual"
				else
					tp_reg_exibe = "Autom&aacute;tico"				 	
				end if							 
				 
				data_split= Split(data_exibe,"/")
				dia_s=data_split(0)
				mes_s=data_split(1)
				ano_s=data_split(2)
				
				
				dia_s=dia_s*1
				mes_s=mes_s*1
				
				if dia_s<10 then
				dia_s="0"&dia_s
				end if
				if mes_s<10 then
				mes_s="0"&mes_s
				end if
			
				da_show=dia_s&"/"&mes_s&"/"&ano_s
			 if da_show = "31/12/9999" then
			 	da_show = ""
			 end if		
		 	if left(nome_exibe,3) = "ZZZ" then
				nome_exibe = replace(nome_exibe,"ZZZ","")
			end if
		 %>                  
						<tr class="<%=cor%>"> 
						  <td width="20"> <input type="checkbox" name="historico" id="historico" class="borda" value="<%=optobr%>"></td>
						  <td width="60" align="center"><%response.Write(ano_exibe)%> </td>
						  <td width="30" align="center"><%response.Write(seq_exibe)%></td>
						  <td width="60" align="center"><%response.Write(matric_exibe)%></td>
						  <td width="280" align="left"><%response.Write(nome_exibe)%></td>
						  <td width="40" align="center"> 
							  <%response.Write(curso_exibe)%>
							  </td>
						  <td width="110" align="center"> 
							<%response.Write(seg_exibe)%>
						  <div align="left"></div></td>
						  <td width="180" align="center"> 
							<%response.Write(escola_exibe)%>
						  </td>
						  <td width="80" align="center"><%response.Write(situac_exibe)%></td>
						  <td width="80" align="center"><%response.Write(tp_reg_exibe)%></td>
						  <td width="60" align="center"> <%response.Write(da_show)%></td>
						</tr>
						<%check = check+1
						intrec = intrec+1
		Rs_ordena.Movenext
		WEND%>
                <tr class="<%=cor%>"> 
                  <td colspan="11">
                    <div align="left"> 
                      <hr width="1000">
                    </div></td>
                </tr>
                <tr class="<%=cor%>"> 
                  <td colspan="11">
	<%if intpagina = 1 and tot_rec<=Rs_ordena.PageSize then
        response.Write("&nbsp;")
    else
        %>                  
                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td class="tb_tit"><div align="center">                          </div></td>
                          </tr>
                          <tr> 
                            <td class="tb_tit"><div align="center"> 
    
                                &nbsp; 
                                <%		  
         if intpagina>1 then
        %>
                                <a href="resumo.asp?obr=<%=obr%>&incl=<%=incl%>&pagina=<%=intpagina-1%>" class="linktres">Anterior</a> 
                                <%
        end if 
        for contapagina=1 to Rs_ordena.PageCount 
                            intpagina=intpagina*1
                            IF contapagina=intpagina then
                                response.Write(contapagina)
                            else
                            %>
                            <a href="resumo.asp?obr=<%=obr%>&incl=<%=incl%>&pagina=<%=contapagina%>" class="linktres"><%response.Write(contapagina)%></a> 
                            <%
                            end if
        next
        if StrComp(intpagina,Rs_ordena.PageCount)<>0 then  
        %>
                                <a href="resumo.asp?obr=<%=obr%>&incl=<%=incl%>&pagina=<%=intpagina + 1%>" class="linktres">Próximo</a> 
                                <%
        end if
        %>
                              </div></td>
                          </tr>
                <tr class="<%=cor%>"> 
                  <td>
                    <div align="left"> 
                      <hr width="1000">
                    </div></td>
                </tr>                             
                        </table>
    <%    
end if	
Rs_ordena.Close




trava="n"


%>
</td>
                </tr>                             
                <tr class="<%=cor%>"> 
                  <td colspan="11"><table width="1000" border="0" align="center" cellspacing="0">
                      <tr> 
                        <td width="200" align="center"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','index.asp?nvg=<%=chave%>');return document.MM_returnValue" value="Voltar">
                            </font></div></td>
                        <td width="200" align="center"><%if trava="n" then%>
                          <div align="center">
                            <input name="submit" type="submit" class="botao_prosseguir" value="Incluir"  >
<!--onClick="MM_goToURL('parent','incluir.asp?cod=<%response.Write(obr)%>&opt=inc');return document.MM_returnValue"-->                            
                          </div>
                        <% end if%></td>
                        <td width="200" align="center"><%if trava="n" then%>
                          <div align="center">
                            <input name="submit" type="submit" class="botao_cancelar" value="Alterar">
                          </div>
                        <% end if%></td>
                        <td width="200" align="center"><%if trava="n" then%>
                          <div align="center">
                            <input name="submit" type="submit" class="botao_excluir" value="Excluir">
                          </div>
                        <% end if%></td>
                        <td width="200" align="center">
                          <div align="center">
                            <input name="submit" type="submit" class="botao_cancelar" value="Imprimir">
                          </div>
                        </td>
                      </tr>
                  </table></td>
                </tr>
                <%

END IF%>
              </table>

              </div></td>
        </tr>
      </table></td>
    </tr>
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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
pasta=arPath(seleciona1)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>