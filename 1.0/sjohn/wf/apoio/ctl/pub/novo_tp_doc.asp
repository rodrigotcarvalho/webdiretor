<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp" -->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->


<%
opt=request.QueryString("opt")
ano_letivo_wf = Session("ano_letivo_wf")
nivel=4
chave=session("chave")
session("chave")=chave
exibe="n"
ano_info=nivel&"-"&chave&"-"&ano_letivo_wf


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
		
	
if transicao = "S" then
	area="wd"
	site_escola="simplynetcloud.hospedagemdesites.ws/wd/"&ambiente_escola&"/wf/apoio/ctl/pub/"
else
	if left(ambiente_escola,5)= "teste" then
		area="wdteste"
		site_escola="www.simplynet.com.br/"&area&"/"&ambiente_escola&"/wf/apoio/ctl/pub"
	else
		area="wd"
		site_escola="www.simplynet.com.br/"&area&"/"&ambiente_escola&"/wf/apoio/ctl/pub"
	end if
end if


call navegacao (CON,chave,nivel)
navega=Session("caminho")	


nome_pasta=""
dia_exp=31
mes_exp=12
ano_exp=ano_letivo_wf
'action="novo_tp_doc.asp?opt=c"
action="grava_novo_tp_doc.asp"
titulo="Criar novo"

	Set RS_doc = Server.CreateObject("ADODB.Recordset")
	SQL_doc = "SELECT MAX(CO_Pasta_Doc) as ultima_pasta FROM TB_Tipo_Pasta_Doc"
	RS_doc.Open SQL_doc, CON0
	if RS_doc.EOF then
		maior_pasta=0
	ELSE
		maior_pasta=RS_doc("ultima_pasta")
		if isnull(maior_pasta) or maior_pasta="" then
			maior_pasta=0		
		end if
	END IF	

	check_s="CHECKED"
	check_n=""	
	disable=""		

if opt="c" then
'	maior_pasta=request.form("maior_pasta")
'	nome_pasta=request.form("nome_pasta")
'	espira=request.form("espira")
'	dia_exp=request.form("dia_exp")
'	mes_exp=request.form("mes_exp")
'	data_exp=dia_exp&"/"&mes_exp&"/"&ano_letivo_wf
'	
'	maior_pasta=maior_pasta*1
'	nova_pasta =maior_pasta+1 
'	
'	if espira="s" then
'		espira=TRUE
'	else	
'		espira=FALSE
'		data_exp=NULL
'	end if	
'
'	Set RS_updt = server.createobject("adodb.recordset")
'	RS_updt.open "TB_Tipo_Pasta_Doc", CON0, 2, 2 'which table do you want open
'	RS_updt.addnew
'	
'		RS_updt("CO_Pasta_Doc")=nova_pasta
'		RS_updt("NO_Pasta") = nome_pasta
'		RS_updt("IN_Expira") = espira
'		RS_updt("DA_Expira") = data_exp
'
'	RS_updt.update
'	set RS_updt=nothing
'	
''RESPONSE.Write("http://"&site_escola&"/sndocs/criapasta.asp?al="&ano_letivo_wf&"&mp="&nova_pasta)
'response.Redirect("http://"&site_escola&"/sndocs/criapasta.asp?al="&ano_letivo_wf&"&mp="&nova_pasta&"&env="&ambiente_escola)
elseif opt="a" then
	tp_doc=request.QueryString("tp")

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Tipo_Pasta_Doc where CO_Pasta_Doc="&tp_doc 
		RS.Open SQL, CON0	
		
	maior_pasta=tp_doc
	nome_pasta=RS("NO_Pasta")
	data_exp=RS("DA_Expira")
	espira=RS("IN_Expira")
	
	IF espira=TRUE THEN
		data_exp=split(data_exp,"/")
	
		dia_exp=data_exp(0)
		mes_exp=data_exp(1)
		ano_exp=data_exp(2)
		check_s="CHECKED"
		check_n=""
		disable=""			
	ELSE
		dia_exp=31
		mes_exp=12
		ano_exp=ano_exp	
		check_s=""
		check_n="CHECKED"	
		disable="disabled"				
	END IF	
	
action="novo_tp_doc.asp?opt=bda"
titulo="Modificar"
elseif opt="bda" then
	tp_doc=request.form("maior_pasta")
	maior_pasta=tp_doc
	nome_pasta=request.form("nome_pasta")
	espira=request.form("espira")
	dia_exp=request.form("dia_exp")
	mes_exp=request.form("mes_exp")
	data_exp=dia_exp&"/"&mes_exp&"/"&ano_letivo_wf
	
	if espira="s" then
		espira=TRUE
		check_s="CHECKED"
		check_n=""		
		disable=""
		sql_atualiza= "UPDATE TB_Tipo_Pasta_Doc SET NO_Pasta='"&nome_pasta&"',DA_Expira=#"&data_exp&"#,IN_Expira="&espira&" WHERE CO_Pasta_Doc = "& tp_doc

	else	
		espira=FALSE
		check_s=""
		check_n="CHECKED"	
		disable="disabled"		
		sql_atualiza= "UPDATE TB_Tipo_Pasta_Doc SET NO_Pasta='"&nome_pasta&"',DA_Expira=NULL,IN_Expira="&espira&" WHERE CO_Pasta_Doc = "& tp_doc		
	end if		

	Set RS_updt2 = CON0.Execute(sql_atualiza)
	
	action="novo_tp_doc.asp?opt=bda"
	titulo="Modificar"
end if

%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="file:../../../../img/mm_menu.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
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

function checksubmit()
{
 if (document.formulario.nome_pasta.value == "")
  {    alert("Por favor digite um nome para o tipo de documento!")
   document.formulario.nome_pasta.focus()
    return false
 } 
  return true
}
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
function habilita_campo(){
		document.getElementById('dia_exp').disabled   = false;
		document.getElementById('mes_exp').disabled   = false;			    
}

function desabilita_campo(){
		document.getElementById('dia_exp').disabled   = true;
		document.getElementById('mes_exp').disabled   = true;	   
}
//-->
</script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../../img/menu_r1_c2_f3.gif','../../../../img/menu_r1_c2_f2.gif','../../../../img/menu_r1_c2_f4.gif','../../../../img/menu_r1_c4_f3.gif','../../../../img/menu_r1_c4_f2.gif','../../../../img/menu_r1_c4_f4.gif','../../../../img/menu_r1_c6_f3.gif','../../../../img/menu_r1_c6_f2.gif','../../../../img/menu_r1_c6_f4.gif','../../../../img/menu_r1_c8_f3.gif','../../../../img/menu_r1_c8_f2.gif','../../../../img/menu_r1_c8_f4.gif','../../../../img/menu_direita_r2_c1_f3.gif','../../../../img/menu_direita_r2_c1_f2.gif','../../../../img/menu_direita_r2_c1_f4.gif','../../../../img/menu_direita_r4_c1_f3.gif','../../../../img/menu_direita_r4_c1_f2.gif','../../../../img/menu_direita_r4_c1_f4.gif','../../../../img/menu_direita_r6_c1_f3.gif','../../../../img/menu_direita_r6_c1_f2.gif','../../../../img/menu_direita_r6_c1_f4.gif')">
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
if opt = "ok" then%>
  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(nivel,74,2,0)
%>
    </td>
    </tr>
      <%
elseif opt = "bda" then%>
  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(nivel,75,2,0)
%>
    </td>
                  </tr>
                  
			  
<% 	end if 

%>                  <tr> 
                    
    <td height="10"> 
      <%	call mensagens(4,69,0,0) 
	  
	  
%>
</td></tr>
<tr>

            <td valign="top"> 		
<FORM name="formulario" METHOD="POST" ACTION="<%response.Write(action)%>" onSubmit="return checksubmit()">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td class="tb_tit"><%response.Write(titulo)%>
Nome da Pasta de Documentos </td>
    </tr>
    <tr>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="45%" class="form_corpo"><div align="right"> Nome da Pasta de Documentos :&nbsp;</div></td>
          <td width="55%"><input name="nome_pasta" type="text" class="textInput" id="nome_pasta" value="<%response.write(nome_pasta)%>" size="50" maxlength="45"></td>
        </tr>
        <tr>
          <td class="form_corpo"><div align="right">Esta pasta expira?&nbsp;</div></td>
          <td><div align="left">
            <table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="25">
                <input name="espira" type="radio" class="borda" id="espira_s" value="s" <%response.Write(check_s)%> onClick="javascript:habilita_campo();" ></td>
                <td width="25" class="form_dado_texto">S</td>
                <td width="25"><input name="espira" type="radio" class="borda" id="espira_n" value="n" <%response.Write(check_n)%> onClick="javascript:desabilita_campo();"></td>
                <td width="25" class="form_dado_texto">N</td>
              </tr>
            </table>
          </div></td>
        </tr>
        <tr>
          <td width="45%" class="form_corpo"><div align="right">Data em que expira:&nbsp;</div></td>
          <td><div align="left"><font class="form_dado_texto"> 
            <select name="dia_exp" id="dia_exp" class="borda" <%response.Write(disable)%>>
              <% 
		 For i =1 to 31
			if i<10 then
			i="0"&i
			end if
			
			i=i*1
			dia_exp=dia_exp*1

			if dia_exp=i then		
							 %>
              <option value="<%response.Write(i)%>" selected> 
                <%response.Write(i)%>
                </option>
              <%else%>
              <option value="<%response.Write(i)%>"> 
                <%response.Write(i)%>
                </option>                             
              <%	end if
		next
							%>                 
              </select>
            / 
            <select name="mes_exp" id="mes_exp" class="borda" <%response.Write(disable)%>>
              <%
							mes_exp=mes_exp*1
								if mes_exp="1" or mes_exp=1 then%>
              <option value="1" selected>janeiro</option>
              <% else%>
              <option value="1">janeiro</option>
              <%end if
								if mes_exp="2" or mes_exp=2 then%>
              <option value="2" selected>fevereiro</option>
              <% else%>
              <option value="2">fevereiro</option>
              <%end if
								if mes_exp="3" or mes_exp=3 then%>
              <option value="3" selected>mar&ccedil;o</option>
              <% else%>
              <option value="3">mar&ccedil;o</option>
              <%end if
								if mes_exp="4" or mes_exp=4 then%>
              <option value="4" selected>abril</option>
              <% else%>
              <option value="4">abril</option>
              <%end if
								if mes_exp="5" or mes_exp=5 then%>
              <option value="5" selected>maio</option>
              <% else%>
              <option value="5">maio</option>
              <%end if
								if mes_exp="6" or mes_exp=6 then%>
              <option value="6" selected>junho</option>
              <% else%>
              <option value="6">junho</option>
              <%end if
								if mes_exp="7" or mes_exp=7 then%>
              <option value="7" selected>julho</option>
              <% else%>
              <option value="7">julho</option>
              <%end if%>
              <%if mes_exp="8" or mes_exp=8 then%>
              <option value="8" selected>agosto</option>
              <% else%>
              <option value="8">agosto</option>
              <%end if
								if mes_exp="9" or mes_exp=9 then%>
              <option value="9" selected>setembro</option>
              <% else%>
              <option value="9">setembro</option>
              <%end if
								if mes_exp="10" or mes_exp=10 then%>
              <option value="10" selected>outubro</option>
              <% else%>
              <option value="10">outubro</option>
              <%end if
								if mes_exp="11" or mes_exp=11 then%>
              <option value="11" selected>novembro</option>
              <% else%>
              <option value="11">novembro</option>
              <%end if
								if mes_exp="12" or mes_exp=12 then%>
              <option value="12" selected>dezembro</option>
              <% else%>
              <option value="12">dezembro</option>
              <%end if%>
              </select>
            / 
            <%response.write(ano_letivo_wf)%></font></div></td>
        </tr>
        <tr>
          <td colspan="2" class="form_corpo"><hr></td>
        </tr>
        <tr>
          <td colspan="2" class="form_corpo"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="33%"><div align="center">
                <input name="SUBMIT5" type=button class="botao_cancelar" onClick="MM_goToURL('parent','index.asp?nvg=<%=chave%>');return document.MM_returnValue" value="Voltar">
              </div></td>
              <td width="34%"><div align="center"></div>
                <div align="center"><input name="maior_pasta" type="hidden" value="<%response.write(maior_pasta)%>"></div></td>
              <td width="33%"><div align="center">
                <input name="SUBMIT2" type=SUBMIT class="botao_prosseguir" value="Salvar">
              </div></td>
              </tr>
          </table></td>
          </tr>
      </table></td>
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
<%
CON.Close
Set CON = Nothing

CON0.Close
Set CON0 = Nothing

CON1.Close
Set CON1 = Nothing

CON_WF.Close
Set CON_WF = Nothing	

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