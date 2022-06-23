<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->

<!--#include file="../../../../inc/funcoes2.asp"-->


<!--#include file="../../../../inc/caminhos.asp"-->

<%
opt=request.QueryString("opt")
pagina=request.QueryString("pagina")
volta=request.QueryString("v")
autoriza = session("autoriza")
session("autoriza")=autoriza
ano_letivo_wf = Session("ano_letivo_wf")
co_usr = session("co_user")
nivel=4
nvg = session("chave")
chave=nvg
session("chave")=chave
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano_info=nivel&"-"&chave&"-"&ano_letivo_wf

if (pagina=1 or pagina="1") and volta="n" then
tp_doc=request.Form("tipo_doc")
dia_de= request.form("dia_de")
mes_de= request.form("mes_de")
dia_ate=request.Form("dia_ate")
mes_ate=request.Form("mes_ate")
unidade=request.Form("unidade")
curso=request.Form("curso")
etapa=request.Form("etapa")
turma=request.Form("turma")
tit=request.Form("tit")


Session("dia_de")=dia_de
Session("mes_de")=mes_de
Session("dia_ate")=dia_ate
Session("mes_ate")=mes_ate
Session("unidade")=unidade
Session("curso")=curso
Session("etapa")=etapa
Session("turma")=turma
Session("tit")=tit
session("tipo_arquivo") =tp_doc


elseif (pagina=1 or pagina="1") and volta="s"then
dia_de= Session("dia_de")
mes_de= Session("mes_de")
dia_ate=Session("dia_ate")
mes_ate=Session("mes_ate")
unidade=Session("unidade")
curso=Session("curso")
etapa=Session("etapa")
turma=Session("turma")
tit=Session("tit")
tp_doc=session("tipo_arquivo")

Session("dia_de")=dia_de
Session("mes_de")=mes_de
Session("dia_ate")=dia_ate
Session("mes_ate")=mes_ate
Session("unidade")=unidade
Session("curso")=curso
Session("etapa")=etapa
Session("turma")=turma
Session("tit")=tit
session("tipo_arquivo") =tp_doc

else
dia_de= Session("dia_de")
mes_de= Session("mes_de")
dia_ate=Session("dia_ate")
mes_ate=Session("mes_ate")
unidade=Session("unidade")
curso=Session("curso")
etapa=Session("etapa")
turma=Session("turma")
tit=Session("tit")
tp_doc=session("tipo_arquivo")

Session("dia_de")=dia_de
Session("mes_de")=mes_de
Session("dia_ate")=dia_ate
Session("mes_ate")=mes_ate
Session("unidade")=unidade
Session("curso")=curso
Session("etapa")=etapa
Session("turma")=turma
Session("tit")=tit
session("tipo_arquivo") =tp_doc

end if
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

if unidade="999990" or unidade="" or isnull(unidade) then
sql_un=""
unidade_nome="Todas"
else
sql_un="(Unidade= '"&unidade&"' OR Unidade is null) AND"

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="&unidade
		RS0.Open SQL0, CON0
		
		unidade_nome = RS0("NO_Unidade")
end if

if curso="999990" or curso="" or isnull(curso) then
sql_cu=""
curso_nome="Todos"
else
sql_cu="(Curso='"&curso&"' OR (Curso  is null)) AND"
		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Curso where CO_Curso='"&curso&"'"
		RS0c.Open SQL0c, CON0
		
		curso_nome = RS0c("NO_Curso")
end if

if etapa="999990" or etapa="" or isnull(etapa) then
sql_et=""
etapa_nome="Todas"
else
sql_et="(Etapa='"&etapa&"' OR (Etapa  is null)) AND"

		Set RS0e = Server.CreateObject("ADODB.Recordset")
		SQL0e = "SELECT * FROM TB_Etapa where CO_Curso='"&curso&"' AND CO_Etapa='"&etapa&"'"
		RS0e.Open SQL0e, CON0
		
		etapa_nome = RS0e("NO_Etapa")
end if

if turma="999990" or turma="" or isnull(turma) then
sql_tu=""
turma_nome="Todas"
else
sql_tu="(Turma='"&turma&"' OR (Turma  is null)) AND "
turma_nome=turma
end if

if tp_doc=0 or  tp_doc="" or isnull(tp_doc) then
sql_tp_doc=""
else
sql_tp_doc="TP_Noticia= "&tp_doc&" AND "
end if

if tit="" or isnull(tit) then
sql_tit=""
tit_nome="Todos"
else
sql_tit="(NT_Titulo like '%"&tit&"%') AND"
tit_nome="Contendo a(s) palavra(s): "&tit
end if

data_de=mes_de&"/"&dia_de&"/"&ano_letivo_wf
data_ate=mes_ate&"/"&dia_ate&"/"&ano_letivo_wf

if volta="s" then
if dia_de<10 then
dia_de="0"&dia_de
end if
end if
if mes_de<10 then
mes_de="0"&mes_de
end if
data_inicio=dia_de&"/"&mes_de&"/"&ano_letivo_wf
if dia_ate<10 then
dia_ate="0"&dia_ate
end if
if mes_ate<10 then
mes_ate="0"&mes_ate
end if
data_fim=dia_ate&"/"&mes_ate&"/"&ano_letivo_wf


		'response.write "SELECT * FROM TB_Documentos where "&sql_tp_doc&" "&sql_un&" "&sql_cu&" "&sql_et&" "&sql_tu&" (DA_Doc BETWEEN #"&data_de&"# AND #"&data_ate&"#) order by DA_Doc Desc"



	




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
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
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

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
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
<%if autoriza=1 then%>
  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(nivel,9701,0,0)
%>
    </td>
                  </tr>
<%end if%> 	  
      <%
if opt = "ok" then%>
  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(nivel,64,2,0)
%>
    </td>
                  </tr>
                  <% 	end if 

%>                  <tr> 
                    
    <td height="10"> 
      <%	call mensagens(4,9707,0,0) 
	  
	  
%>
</td></tr>
<tr>

            <td valign="top"> 
			
	
			
<FORM name="formulario" METHOD="POST" ACTION="confirma.asp">
                
        <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td width="653" height="15" class="tb_tit">Crit&eacute;rios da pesquisa 
              <input name="co_grupo" type="hidden" id="co_grupo" value="<% = co_grupo %>"></td>
          </tr>
          <tr> 
            <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <%
		Set RS_doc = Server.CreateObject("ADODB.Recordset")
'		SQL_doc = "SELECT * FROM TB_Documentos where "&sql_tp_doc&" "&sql_un&" "&sql_cu&" "&sql_et&" "&sql_tu&" OR  ((Unidade Is Null) AND (Curso Is Null) AND  (Etapa Is Null) AND (Turma Is Null))) AND (DA_Doc BETWEEN #"&data_de&"# AND #"&data_ate&"#) order by DA_Doc Desc"
SQL_doc = "SELECT * FROM TB_Noticias where "&sql_tp_doc&sql_un&sql_cu&sql_et&sql_tu&sql_tit&"(NT_DT_Pb BETWEEN #"&data_de&"# AND #"&data_ate&"#) order by NT_DT_Pb Desc"
		RS_doc.Open SQL_doc, CON_WF, 3, 3

    if cint(Request.QueryString("pagina"))<1 then
	intpagina = 1
    else
		if cint(Request.QueryString("pagina"))>RS_doc.PageCount then  
	    intpagina = RS_doc.PageCount
        else
    	intpagina = Request.QueryString("pagina")
		end if
    end if   


	
 RS_doc.PageSize = 30
 
if Request.QueryString("pagina")="" then
      intpagina = 1
else
    if cint(Request.QueryString("pagina"))<1 then
	intpagina = 1
    else
		if cint(Request.QueryString("pagina"))>RS_doc.PageCount then  
	    intpagina = RS_doc.PageCount
        else
    	intpagina = Request.QueryString("pagina")
		end if
    end if   
 end if   

		

%>
                <tr class="<%response.write(cor)%>"> 
                  <td colspan="10" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="tb_subtit"><div align="center">Tipo de Not&iacute;cia</div></td>
                        <td colspan="2" class="tb_subtit"><div align="center">Per&iacute;odo 
                            da Publica&ccedil;&atilde;o</div></td>
                        <td class="tb_subtit"><div align="center">T&iacute;tulo 
                            da Not&iacute;cia</div></td>
                      </tr>
                      <tr> 
                        <td><div align="center"><font class="form_dado_texto"> 
                            <%
if tp_doc=0 or  tp_doc="" or isnull(tp_doc) then
tipo_doc_nome= "Todos"
else


		Set RS1n = Server.CreateObject("ADODB.Recordset")
		SQL1n = "SELECT * FROM TB_Tipo_Noticias where TP_Noticia="&tp_doc
		RS1n.Open SQL1n, CON0


tipo_doc_nome=RS1n("TX_Descricao")
end if
response.Write(tipo_doc_nome)
%>
                            </font></div></td>
                        <td colspan="2"><div align="center"><font class="form_dado_texto"> 
                            <%response.Write(data_inicio)%>
                            at&eacute; 
                            <%response.Write(data_fim)%>
                            </font></div></td>
                        <td><div align="center"><font class="form_dado_texto"> 
                            <%response.Write(tit_nome)%>
                            </font></div></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td colspan="2">&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="250" class="tb_subtit"> <div align="center">UNIDADE 
                          </div></td>
                        <td width="250" class="tb_subtit"> <div align="center">CURSO 
                          </div></td>
                        <td width="250" class="tb_subtit"> <div align="center">ETAPA 
                          </div></td>
                        <td width="250" class="tb_subtit"> <div align="center">TURMA 
                          </div></td>
                      </tr>
                      <tr> 
                        <td width="250"> <div align="center"> <font class="form_dado_texto"> 
                            <%response.Write(unidade_nome)%>
                            </font> </div></td>
                        <td width="250"> <div align="center"> <font class="form_dado_texto"> 
                            <%response.Write(curso_nome)%>
                            </font> </div></td>
                        <td width="250"> <div align="center"> <font class="form_dado_texto"> 
                            <%response.Write(etapa_nome)%>
                            </font> </div></td>
                        <td width="250"> <div align="center"> <font class="form_dado_texto"> 
                            <%response.Write(turma_nome)%>
                            </font> </div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr class="<%response.write(cor)%>"> 
                  <td colspan="10" valign="top" ><hr width="1000"></td>
                </tr>
                <tr class="<%response.write(cor)%>"> 
                  <td colspan="10" valign="top" class="tb_tit">Not&iacute;cias</td>
                </tr>
                <%
		if RS_doc.EOF then
intpagina=1
sem_link=1
		%>
                <tr class="<%response.write(cor)%>"> 
                  <td colspan="10" valign="top"> <div align="center"><font class="form_corpo"> 
                      <%response.Write("Não existem documentos para os critérios informados!")%>
                      </font></div></td>
                </tr>
                <%else 
				
sem_link=0
    RS_doc.AbsolutePage = intpagina
    intrec = 0				
				%>
                <tr> 
                  <td width="20" class="tb_subtit"> <div align="center"> 
                      <input type="checkbox" name="todos" class="borda" value="" onClick="this.value=check(this.form.doc)">
                    </div></td>
                  <td width="100" class="tb_subtit"> <div align="center">Tipo 
                      de Not&iacute;cia</div></td>
                  <td width="100" class="tb_subtit"> <div align="center">Publica&ccedil;&atilde;o</div></td>
                  <td width="100" class="tb_subtit"> <div align="center">Vig&ecirc;ncia</div></td>
                  <td width="420" class="tb_subtit"> 
                    <div align="left">&nbsp;&nbsp;T&iacute;tulo 
                      da Not&iacute;cia</div></td>
                  <td width="60" class="tb_subtit"> 
                    <div align="center">Un</div></td>
                  <td width="60" class="tb_subtit"> 
                    <div align="center">Curso 
                    </div></td>
                  <td width="60" class="tb_subtit"> 
                    <div align="center">Etapa</div></td>
                  <td width="60" class="tb_subtit"> 
                    <div align="center">Turma</div></td>
                </tr>
                <tr class="<%response.write(cor)%>"> 
                  <td colspan="10"><hr width="1000"></td>
                </tr>
                <%				
check=2
while intrec<RS_doc.PageSize and not RS_doc.eof

 if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if
co_doc=RS_doc("NT_Codigo") 
tipo_doc =RS_doc("TP_Noticia") 
tit1=RS_doc("NT_Titulo")
da_vig=RS_doc("NT_DT_Vg")
da_doc=RS_doc("NT_DT_Pb")
unidade=RS_doc("Unidade")
curso=RS_doc("Curso")
etapa=RS_doc("Etapa")
turma=RS_doc("Turma")



if unidade="" or isnull(unidade) then
no_unidade=""
else
 		Set RSnoun = Server.CreateObject("ADODB.Recordset")
		SQLnoun = "SELECT * FROM TB_Unidade Where NU_Unidade="&unidade
		RSnoun.Open SQLnoun, CON0
		
no_unidade=RSnoun("NO_Abr")
end if		

if curso="" or isnull(curso) then
no_curso=""
else



 		Set RSnocu = Server.CreateObject("ADODB.Recordset")
		SQLnocu = "SELECT * FROM TB_Curso Where CO_Curso='"&curso&"'"
		RSnocu.Open SQLnocu, CON0
		
no_curso=RSnocu("NO_Abreviado_Curso")		
end if

if etapa="" or isnull(etapa) then
no_etapa=""
else
 		Set RSnoet = Server.CreateObject("ADODB.Recordset")
		SQLnoet = "SELECT * FROM TB_Etapa Where CO_Curso='"&curso&"' AND CO_Etapa='"&etapa&"'"
		RSnoet.Open SQLnoet, CON0
		
no_etapa=RSnoet("NO_Etapa")		
end if


		Set RS1n = Server.CreateObject("ADODB.Recordset")
		SQL1n = "SELECT * FROM TB_Tipo_Noticias where TP_Noticia="&tipo_doc
		RS1n.Open SQL1n, CON0


tipo_doc_nome=RS1n("TX_Descricao")





data_split= Split(da_doc,"/")
dia=data_split(0)
mes=data_split(1)
ano=data_split(2)


dia=dia*1
mes=mes*1
hora=hora*1
min=min*1

if dia<10 then
dia="0"&dia
end if
if mes<10 then
mes="0"&mes
end if

da_show=dia&"/"&mes&"/"&ano


if da_vig="" or isnull(da_vig) then
da_show_vig=""
else
data_vig_split= Split(da_vig,"/")
dia_vig=data_vig_split(0)
mes_vig=data_vig_split(1)
ano_vig=data_vig_split(2)


dia_vig=dia_vig*1
mes_vig=mes_vig*1


if dia_vig<10 then
dia_vig="0"&dia_vig
end if

if mes_vig<10 then
mes_vig="0"&mes_vig
end if

da_show_vig=dia_vig&"/"&mes_vig&"/"&ano_vig
end if
%>
                <tr class="<%response.write(cor)%>"> 
                  <td width="20"> <div align="center"><font class="form_dado_texto"> 
                      <input name="doc" type="checkbox" class="borda" value="<%=co_doc%>">
                      </font></div></td>
                  <td width="100"> <div align="center"> 
                      <%response.Write(tipo_doc_nome)%>
                    </div></td>
                  <td width="100"> <div align="center"> 
                      <%response.Write(da_show)%>
                    </div></td>
                  <td width="100"> <div align="center"> 
                      <%response.Write(da_show_vig)%>
                    </div></td>
                  <td width="420">&nbsp;&nbsp;<a href="alterar.asp?c=<%=co_doc%>" class="linkum"> 
                    <%response.Write(tit1)%>
                    </a> 
                    <div align="left"></div></td>
                  <td width="60"> 
                    <div align="center"> 
                      <%response.Write(no_unidade)%>
                    </div></td>
                  <td width="60"> 
                    <div align="center"> 
                      <%response.Write(no_curso)%>
                    </div></td>
                  <td width="60"> 
                    <div align="center"> 
                      <%response.Write(no_etapa)%>
                    </div></td>
                  <td width="60"> 
                    <div align="center"> 
                      <%response.Write(turma)%>
                    </div></td>
                </tr>
                <%
intrec = intrec + 1
check=check+1

RS_doc.movenext

wend
end if%>

                <tr> 
                  <td class="tb_tit" colspan="10"><div align="center"> 
                      <%
if sem_link=0 then
	%>
                      &nbsp; 
                      <%		  
			    if intpagina>1 then
    %>
                      <a href="docs.asp?pagina=<%=intpagina-1%>" class="linktres">Anterior</a> 
                      <%
    end if
		for contapagina=1 to RS_doc.PageCount 
						pagina=pagina*1
						IF contapagina=pagina then
						response.Write(contapagina)
						else
						%>
						<a href="docs.asp?pagina=<%=contapagina%>" class="linktres"><%response.Write(contapagina)%></a> 
						<%
						end if
						next
    if StrComp(intpagina,RS_doc.PageCount)<>0 then  
    %>
                      <a href="docs.asp?pagina=<%=intpagina + 1%>" class="linktres">Próximo</a> 
                      <%
    end if
else	
	%>
                      &nbsp; 
                      <%
end if	
    RS_doc.close
    Set RS_doc = Nothing
    %>
                    </div></td>
                </tr>
                <tr> 
                  <td colspan="10"><hr width="1000"></td>
                </tr>
                <tr> 
                  <td colspan="10"><div align="center"> 
                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td width="100%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                              <tr> 
                                <td width="33%"> <div align="center"> 
                                    <input name="SUBMIT5" type=button class="botao_cancelar" onClick="MM_goToURL('parent','index.asp?nvg=<%=nvg%>');return document.MM_returnValue" value="Voltar">
                                  </div></td>
                                <td width="34%"> <div align="center"> 
								  <%if autoriza=1 then
								  else
								  %>
                                    <input name="SUBMIT3" type=submit class="botao_excluir" value="Excluir">
								<%end if%>	
                                  </div></td>
                                <td width="33%"> <div align="center"> 
								  <%if autoriza=1 then
								  else
								  %>								
                                    <input name="SUBMIT2" type=button class="botao_prosseguir" onClick="MM_goToURL('parent','incluir.asp');return document.MM_returnValue" value="Incluir">
                                  <%end if%>
								  </div></td>
                              </tr>
                              <tr> 
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                              </tr>
                            </table></td>
                        </tr>
                      </table>
                    </div></td>
                </tr>
              </table>
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
<%call GravaLog (nvg,"")%>
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