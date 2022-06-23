<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<% 

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
ori = request.QueryString("or")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

trava=session("trava")
exclui_entrevista=request.QueryString("opt")


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

		Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4 = "DBQ="& CAMINHO_e & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4	
		
		
		Set CONp = Server.CreateObject("ADODB.Connection") 
		ABRIRp = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONp.Open ABRIRp		
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		




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
			  
        <form action="bd.asp?opt=exc" method="post" name="busca" id="busca">
          <tr>
      <td valign="top"> 
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>            <tr> 
              
    <td width="766" height="10" colspan="4" valign="top"> 
      <%call mensagens(nivel,326,0,0) %>
    </td>
			  </tr>
          <tr> 
            <td height="10" class="tb_tit"
>Entrevistas a serem exclu&iacute;das</td>
          </tr>
          <tr> 
            <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
                <tr class="tb_subtit"> 
                  <td width="30" height="10"><div align="center"></div></td>
                  <td width="90" align="center">Data / Hora<font class="form_dado_texto">
                  </font></td>
                  <td width="70" align="center">Matr&iacute;cula<font class="form_dado_texto">
                  </font></td>
                  <td width="288"><div align="left">Nome do Aluno<font class="form_dado_texto">
                  </font></div></td>
                  <td width="108" align="center">Tipo</td>
                  <td width="122"><div align="center">Participantes</div></td>
                  <td width="160"><div align="center">Atendido por</div></td>
                  <td width="159"><div align="center">Status<font class="form_dado_texto">
                    <input name="exclui_entrevista" type="hidden" class="textInput" id="exclui_entrevista"  value="<%response.Write(exclui_entrevista)%>" size="75" maxlength="50">
                  </font></div></td>
                </tr>
                <%
'response.Write(">>"&exclui_entrevista)				
check = 2				
vertorExclui = split(exclui_entrevista,", ")
conta_ocorr=0
for i =0 to ubound(vertorExclui)

exclui = split(vertorExclui(i),"?")

'obr=cod&"?"&da_entrevista&"?"&ho_entrevista&"?"&co_entrevista
cod = exclui(0)
da_entrevista= exclui(1)
ho_entrevista= exclui(2)

				
dados_data=split(da_entrevista,"/")
dia=dados_data(0)
mes=dados_data(1)
ano=dados_data(2)

dados_hora=split(ho_entrevista,":")
h=dados_hora(0)
m=dados_hora(1)


da_entrevista_cons=mes&"/"&dia&"/"&ano

h=h*1
m=m*1


ho_entrevista_cons=h&":"&m


	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod
	RS.Open SQL, CON1

	no_aluno = RS("NO_Aluno")	
'response.Write(ho_entrevista_cons)				
		Set RSo = Server.CreateObject("ADODB.Recordset")
		SQLo = "SELECT * FROM TB_Entrevistas WHERE CO_Matricula ="& cod&" AND (DA_Entrevista=#"&da_entrevista_cons&"# AND mid(HO_Entrevista,1,16)=#12/30/1899 "&ho_entrevista_cons&"#)" 
		RSo.Open SQLo, CON4

  if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if 
  

da_entrevista=RSo("DA_entrevista")
ho_entrevista=RSo("HO_Entrevista")
tp_entrevista=RSo("TP_Entrevista")
partic_entrevista=RSo("NO_Participantes")
st_entrevista=RSo("ST_Entrevista")
ag_entrevista=RSo("CO_Agendado_com")
ob_entrevista=RSo("TX_Observa")
cu_entrevista=RSo("CO_Usuario")

dados_hora=split(ho_entrevista,":")
h=dados_hora(0)
m=dados_hora(1)


h=h*1
m=m*1

if h<10 then
h="0"&h
end if

if m<10 then
m="0"&m
end if

ho_entrevista=h&":"&m



 
Set RSto = Server.CreateObject("ADODB.Recordset")
SQLto = "SELECT * FROM TB_Tipo_Entrevista WHERE TP_Entrevista ="& tp_entrevista
RSto.Open SQLto, CON0

if RSto.EOF then
	no_entrevista=""
else
	no_entrevista=RSto("TX_Descricao")
end if
		
if ag_entrevista="" or isnull(ag_entrevista) then
else
	Set RSu = Server.CreateObject("ADODB.Recordset")
	SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& ag_entrevista
	RSu.Open SQLu, CON

	IF RSu.EOF then
		no_atendido=""
	else
		no_atendido=RSu("NO_Usuario")
	end if
	
end if
opt=cod&"?"&da_entrevista&"?"&ho_entrevista

%>
                <tr class="<%=cor%>"> 
                  <td width="30">&nbsp;</td>
                  <td width="90"> <div align="center"> 
                      <%response.Write(da_entrevista&", "&ho_entrevista)%>
                    </div></td>
                  <td width="70"><div align="center"> 
                      <%response.Write(cod)%>
                    </div></td>
                  <td width="288"> 
                      <%response.Write(no_aluno)%>
                   </td>
                  <td width="108"><div align="center"> 
                      <%response.Write(no_entrevista)%>
                    </div></td>
                  <td width="122"><div align="center"> 
                      <%response.Write(partic_entrevista)%>
                    </div></td>
                  <td width="160" align="center"><%response.Write(no_atendido)%></td>
                  <td width="159"> <div align="center"> 
                      <%
					  		Select case st_entrevista		
								case 1
								nome_status="Atendida"
								
								case 2
								nome_status="Cancelada"
								
								case 3
								nome_status="Pendente"	
							end select	
					  
					  
					  response.Write(nome_status)%>
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
                    <td width="391"> <div align="center"> 
                        <input name="alterar" type="submit" class="botao_cancelar" id="alterar" onClick="MM_goToURL('parent','resumo.asp?or=2');return document.MM_returnValue" value="Voltar">
                      </div></td>
                    <td width="391">&nbsp;</td>
                    <td width="218"> <div align="left"> 
                        <input name="Submit" type="submit" class="botao_prosseguir" value="Confirmar">
                      </div></td>
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