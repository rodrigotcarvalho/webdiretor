<%On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/connect_pr.asp"-->
<!--#include file="../../../../inc/connect_l.asp"-->
<!--#include file="../../../../inc/connect_wf.asp"-->

<% 

nivel=4
permissao = session("permissao") 
ano_letivo_wf = session("ano_letivo_wf")  
sistema_local=session("sistema_local")
nvg = session("chave")
chave=nvg
session("chave")=chave
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)


dia_de= Session("dia_de")
mes_de= Session("mes_de")
dia_ate=Session("dia_ate")
mes_ate=Session("mes_ate")
unidade=Session("unidade")
curso=Session("curso")
etapa=Session("etapa")
turma=Session("turma")
tit=Session("tit")
check_status=Session("check_status")
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
Session("check_status")=check_status
session("tipo_arquivo") =tp_doc



trava=session("trava")
exclui_doc=request.form("doc")


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")		


		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		


    	Set CON_WF = Server.CreateObject("ADODB.Connection") 
		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_WF.Open ABRIR_WF

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
  {    alert("Por favor digite SOMENTE uma op��o de busca!")
    document.busca.busca1.focus()
    return false
  }
    if (document.busca.busca1.value == "" && document.busca.busca2.value == "")
  {    alert("Por favor digite uma op��o de busca!")
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
}

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
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
			  
        <form action="bd.asp?opt=e" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
          <tr>
      <td valign="top"> 
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
          <tr> 
            <td width="766" height="10" colspan="4" valign="top"> 
              <%call mensagens(nivel,65,0,0) %>
            </td>
          </tr>
          <tr> 
            <td height="10" class="tb_tit"
>Eventos a serem exclu&iacute;dos</td>
          </tr>
          <tr> 
            <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="20" class="tb_subtit"> <div align="center"><font class="form_dado_texto"> 
                      <input name="exclui_doc" type="hidden" id="exclui_doc" value="<%=exclui_doc%>">
                      </font> </div></td>
                  <td width="70" class="tb_subtit"> <div align="center">In&iacute;cio</div></td>
                  <td width="70" class="tb_subtit"> <div align="center">Fim</div></td>
                  <td width="200" class="tb_subtit"> <div align="center">Tipo 
                      de Calend&aacute;rio</div></td>
                  <td width="200" class="tb_subtit"> <div align="center">Tipo 
                      de Evento</div></td>
                  <td width="220" class="tb_subtit"> <div align="left">&nbsp; 
                      Nome do Evento</div></td>
                  <td width="40" class="tb_subtit"> <div align="center">Repete</div></td>
                  <td width="40" class="tb_subtit"> <div align="center">Un</div></td>
                  <td width="40" class="tb_subtit"> <div align="center">Curso 
                    </div></td>
                  <td width="60" class="tb_subtit"> <div align="center">Etapa</div></td>
                  <td width="40" class="tb_subtit"> <div align="center">Turma</div></td>
                </tr>
                <%
'response.Write(">>"&exclui_ocorrencia)				
check = 2				
vertorExclui = split(exclui_doc,", ")
conta_ocorr=0
for i =0 to ubound(vertorExclui)

co_doc = vertorExclui(i)
		
		Set RS_doc = Server.CreateObject("ADODB.Recordset")
SQL_doc = "SELECT * FROM TB_Agenda where NE_Codigo="&co_doc
		RS_doc.Open SQL_doc, CON_WF

  if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if 
  
tipo_doc =RS_doc("TP_Calen") 
tit1=RS_doc("NO_Evento")
tipo_eve=RS_doc("TP_Evento")
da_in=RS_doc("EV_DT_IN")
da_fim=RS_doc("EV_DT_FI")
repete=RS_doc("EV_Repete")
unidade=RS_doc("Unidade")
curso=RS_doc("Curso")
etapa=RS_doc("Etapa")
turma=RS_doc("Turma")

Select Case repete
case False
repete_nome="N"
case true
repete_nome="S"
end select



 		Set RScal = Server.CreateObject("ADODB.Recordset")
		SQLcal = "SELECT * FROM TB_Tipo_Calendario Where TP_Calen='"&tipo_doc&"'"
		RScal.Open SQLcal, CON0

tipo_doc_nome=RScal("TX_Descricao")

 		Set RSeve = Server.CreateObject("ADODB.Recordset")
		SQLeve = "SELECT * FROM TB_Tipo_Eventos Where TP_Evento="&tipo_eve
		RSeve.Open SQLeve, CON0

tipo_eve_nome=RScal("TX_Descricao")

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



 		Set RScal = Server.CreateObject("ADODB.Recordset")
		SQLcal = "SELECT * FROM TB_Tipo_Calendario Where TP_Calen='"&tipo_doc&"'"
		RScal.Open SQLcal, CON0

tipo_doc_nome=RScal("TX_Descricao")

 		Set RSeve = Server.CreateObject("ADODB.Recordset")
		SQLeve = "SELECT * FROM TB_Tipo_Eventos Where TP_Evento="&tipo_eve
		RSeve.Open SQLeve, CON0

tipo_eve_nome=RSeve("TX_Descricao")


data_split= Split(da_in,"/")
dia=data_split(0)
mes=data_split(1)
ano=data_split(2)


dia=dia*1
mes=mes*1


if dia<10 then
dia="0"&dia
end if
if mes<10 then
mes="0"&mes
end if

da_show=dia&"/"&mes&"/"&ano

data_split_fim= Split(da_fim,"/")
dia_fim=data_split_fim(0)
mes_fim=data_split_fim(1)
ano_fim=data_split_fim(2)


dia_fim=dia_fim*1
mes_fim=mes_fim*1


if dia_fim<10 then
dia_fim="0"&dia_fim
end if
if mes_fim<10 then
mes_fim="0"&mes_fim
end if

da_show_fim=dia_fim&"/"&mes_fim&"/"&ano_fim
%>
                <tr class="<%response.write(cor)%>"> 
                  <td width="20"> <div align="center"><font class="form_dado_texto"> 
                      </font></div></td>
                  <td valign="top"> <div align="center"> 
                      <%response.Write(da_show)%>
                    </div></td>
                  <td valign="top"> <div align="center"> 
                      <%response.Write(da_show_fim)%>
                    </div></td>
                  <td valign="top"> <div align="center"> 
                      <%response.Write(tipo_doc_nome)%>
                    </div></td>
                  <td valign="top"> <div align="center"> 
                      <%response.Write(tipo_eve_nome)%>
                    </div></td>
                  <td valign="top"> <div align="left"> &nbsp; <a href="alterar.asp?c=<%=co_doc%>" class="linkum"> 
                      <%response.Write(tit1)%>
                      </a> </div></td>
                  <td valign="top"> <div align="center"> 
                      <%response.Write(repete_nome)%>
                    </div></td>
                  <td valign="top"> <div align="center"> 
                      <%response.Write(no_unidade)%>
                    </div></td>
                  <td valign="top"> <div align="center"> 
                      <%response.Write(no_curso)%>
                    </div></td>
                  <td valign="top"> <div align="center"> 
                      <%response.Write(no_etapa)%>
                    </div></td>
                  <td valign="top"> <div align="center"> 
                      <%response.Write(turma)%>
                    </div></td>
                </tr>
                <%check = check+1
next
%>
              </table></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></div></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td><hr></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td><div align="center"> 
                <table width="1000" border="0" align="center" cellspacing="0">
                  <tr> 
                    <td width="1000"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td width="25%"> <div align="center"> 
                              <input name="SUBMIT52" type=button class="botao_cancelar" onClick="MM_goToURL('parent','docs.asp?pagina=1&v=s');return document.MM_returnValue" value="Voltar">
                            </div></td>
                          <td width="25%"> <div align="center"> </div></td>
                          <td width="25%"> <div align="center"> </div></td>
                          <td width="25%"> <div align="center"> 
                              <input name="Submit2" type="submit" class="botao_prosseguir" value="Confirmar">
                            </div></td>
                        </tr>
                      </table></td>
                  </tr>
                </table>
                <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> </font></div></td>
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