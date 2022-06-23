<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<% 
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
session("cod_url") = ""	
obr=request.QueryString("obr")
opt=request.QueryString("opt")


chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

vetor_cod_cons = session("vetor_cod_cons")
session("vetor_cod_cons") = vetor_cod_cons


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
	ABRIR7 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON7.Open ABRIR7
			
call VerificaAcesso (CON,chave,nivel)
autoriza=Session("autoriza")

call navegacao (CON,chave,nivel)
navega=Session("caminho")

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

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
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
//Nessa função o formulário a ser enviadopor essa função é o segundo
//function submitfuncao()  
//{
//   var f=document.forms[3]; 
//      f.submit(); 
//}
function submitfuncao()  
{
   var f=document.forms[4]; 
      f.submit(); 
}
//-->
</script>

</head>

<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<%call cabecalho(nivel)
%>
<table width="1000" height="685" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
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
      <%call mensagens(nivel,713,0,0) %>
    </td>
			  </tr>			  

 <tr class="tb_corpo"> 
                  
    <td height="10" colspan="5" class="tb_tit">Alunos que terão o bonus excluído</td>
                </tr>
<form name="seleciona_alunos" method="post" action="bd.asp?opt=exc">                  
    <tr> 
<%
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = 	"Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada, TB_Alunos.NO_Aluno, TB_Matriculas.NU_Unidade,TB_Matriculas.CO_Curso,TB_Matriculas.CO_Etapa,TB_Matriculas.CO_Turma from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Situacao='C' AND TB_Matriculas.CO_Matricula in ("&vetor_cod_cons&") AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula order by TB_Matriculas.NU_Unidade,TB_Matriculas.CO_Curso,TB_Matriculas.CO_Etapa,TB_Matriculas.CO_Turma,TB_Matriculas.NU_Chamada"
		RS.Open SQL, CON1
%>                  
    <td colspan="5" valign="top"> 
    <table width="1000" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="80" class="tb_subtit"><div align="center">Matr&iacute;cula</div></td>
    <td width="400" class="tb_subtit">&nbsp;Nome<input name="obr" type="hidden" value="<%response.write(obr)%>">
<input name="alunos" type="hidden" value="<%response.write(obr)%>">    </td>
    <td width="100" height="10" class="tb_subtit"><div align="center">Unidade</div></td>
    <td width="100" height="10" class="tb_subtit"><div align="center">Curso</div></td>
    <td width="100" height="10" class="tb_subtit"><div align="center"> Etapa</div></td>
    <td width="100" height="10" class="tb_subtit"><div align="center">Turma </div></td>
    <td width="100" class="tb_subtit"><div align="center">Bonus</div></td>
    </tr>
        <%
check=0		
WHile Not RS.EOF
	nu_matricula = RS("CO_Matricula")
	no_aluno= RS("NO_Aluno")			
	nu_chamada = RS("NU_Chamada")
	unidade_aluno = RS("NU_Unidade")
	curso_aluno = RS("CO_Curso")
	co_etapa_aluno = RS("CO_Etapa")
	turma_aluno = RS("CO_Turma")
	
call GeraNomes("PORT",unidade_aluno,curso_aluno,co_etapa_aluno,CON0)
no_unidade = session("no_unidades")
no_curso = session("no_grau")
no_etapa = session("no_serie")	
	
	 if check mod 2 =0 then
		cor = "tb_fundo_linha_par" 
	 else 
		cor ="tb_fundo_linha_impar"
	 end if 
	 
		Set RSB = Server.CreateObject("ADODB.Recordset")
		SQLB = 	"Select * from TB_Bonus_Media_Anual WHERE CO_Matricula = "&nu_matricula
		RSB.Open SQLB, CON7 	 
		
		if RSB.eof then	 
 		
		else
		
		val_bonus = RSB("bonus")
%>		
  <tr class="<%=cor%>">
    <td width="80"><div align="center"><%response.Write(nu_matricula)%></div></td>
    <td width="400">&nbsp;<%response.Write(no_aluno)%></td>
    <td width="100"><div align="center"><%response.Write(no_unidade)%></div></td>
    <td width="100"><div align="center"><%response.Write(no_curso)%></div></td>
    <td width="100"><div align="center"><%response.Write(no_etapa)%></div></td>
    <td width="100"><div align="center"><%response.Write(turma_aluno)%></div></td>
    <td width="100"><div align="center">
      <%response.Write(val_bonus)%>
    </div></td>
    </tr>
	
 <% end if
check=check+1
RS.Movenext
Wend

if check < 1 then

%>		
  <tr class="<%=cor%>">
    <td width="20" colspan = "7" align="center">Não foram encontrados alunos com bonus de média anual para a combinação de unidade, curso, etapa e turma informada
  </tr>
	
 <% end if

%>
  <tr class="tb_fundo_linha_par">
    <td colspan="7">
    <table width="1000" border="0" align="center" cellspacing="0">
                      <tr>
                        <td colspan="3"><hr></td>
                      </tr>
                      <tr> 
                        <td width="375"><div align="center">
                          <!--<input name="Submit" type="submit" class="botao_prosseguir" value="Incluir">-->
                          <input name="Submit2" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','select_alunos.asp?vt=s&obr=<%response.Write(obr)%>');return document.MM_returnValue" value="Voltar">
                        </div></td>
                        <td width="375">&nbsp;</td>
                        <td width="250"> 
                          <div align="center"> 
                            <!--<input name="Submit" type="submit" class="botao_prosseguir" value="Incluir">-->
                            <input name="Submit" type="submit" class="botao_prosseguir" value="Confirmar">
                          </div>
                        </td>
                </tr>
    </table>
    </td>
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