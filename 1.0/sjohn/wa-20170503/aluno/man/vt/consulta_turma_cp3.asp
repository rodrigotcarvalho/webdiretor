<%On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes3.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->

<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->

<%opt = REQUEST.QueryString("opt")
obr = request.QueryString("o")
nivel=4

opt=request.QueryString("opt")

autoriza=Session("autoriza")
Session("autoriza")=autoriza

permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

if opt="direto" then

curso = request.querystring("curso")
unidade = request.querystring("unidade")
grade = request.querystring("grade")
turma = "sem turma"

elseif opt= "vt" then
dados= split(obr, "_" )
unidade= dados(0)
curso= dados(1)
co_etapa= dados(2)
turma= dados(3)


else

curso = request.Form("curso")
unidade = request.Form("unidade")
co_etapa = request.Form("etapa")
turma = request.Form("turma")
end if
ano_letivo = session("ano_letivo")
obr=unidade&"_"&curso&"_"&co_etapa&"_"&turma
if co_etapa = "f0"then
co_etapa=0
elseif co_etapa = "f1" or co_etapa = "m1"then
co_etapa=1
elseif co_etapa = "f2" or co_etapa = "m2"then
co_etapa = 2
elseif co_etapa = "f3" or co_etapa = "m3"then
co_etapa = 3
elseif co_etapa = "f4" then
co_etapa = 4
elseif co_etapa = "f5" then
co_etapa = 5
elseif co_etapa = "f6" then
co_etapa = 6
elseif co_etapa = "f7" then
co_etapa = 7
elseif co_etapa = "f8" then
co_etapa = 8
elseif co_etapa = "f55" then
co_etapa = 55
elseif co_etapa = "f66" then
co_etapa = 66
elseif co_etapa = "f77" then
co_etapa = 77
end if

if co_etapa="999999" then 
response.Redirect("tabelas2.asp?opt=err2&or=02&curso="&curso&"&unidade="&unidade&"")
else

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Abreviado_Curso")

		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' and CO_Curso ='"& curso &"'"  
		RS3.Open SQL3, CON0

 call navegacao (CON,chave,nivel)
navega=Session("caminho")
		
if RS3.EOF THEN
no_etapa="sem etapa"
else
no_etapa=RS3("NO_Etapa")
end if
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
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
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
	call mensagens(nivel,305,0,0) 
		call ultimo(0) 
%>
    </td>
                </tr>
                <tr> 
                  
    <td valign="top"> 
      <form name="inclusao" method="post" action="grade_cp4i.asp" onSubmit="return checksubmit()">
                <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
                  <tr class="tb_tit"
> 
                    <td width="653" height="15" class="tb_tit"
>Segmento </td>
                  </tr>
                  <tr> 
                    <td><table width="1000" border="0" cellspacing="0">
                <tr> 
                  <td width="8"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                  <td class="tb_subtit"> <div align="center">UNIDADE </div></td>
                  <td class="tb_subtit"> <div align="center">CURSO </div></td>
                  <td class="tb_subtit"> <div align="center">ETAPA </div></td>
                  <td class="tb_subtit"> <div align="center">TURMA </div></td>
                </tr>
                <tr> 
                  <td width="8"> </td>
                  <td> <div align="center"> <font class="form_dado_texto">  
                      <%response.Write(no_unidade)%>
                      </font></div></td>
                  <td> <div align="center"> <font class="form_dado_texto">  
                      <%
response.Write(no_curso)%>
                      </font></div></td>
                  <td><div align="center"> <font class="form_dado_texto">  
                      <%


response.Write(no_etapa)%>
                      <input name="etapa" type="hidden" id="etapa" value="<% = co_etapa %>">
                      </font></div></td>
                  <td> <div align="center"> <font class="form_dado_texto">  
                      <%response.Write(turma)%>
                      <input name="turma" type="hidden" id="turma" value="<% = turma%>">
                      </font></div></td>
                </tr>
              </table></td>
                  </tr>
                  <tr bgcolor="#FFFFFF"> 
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>
                  <tr> 
                    <td> <table width="1000" border="0" cellspacing="0">
                <tr> 
                  <td width="10"> </td>
                  <td width="70" class="tb_subtit"> 
                    <div align="center">Matricula 
                    </div></td>
                  <td width="70" class="tb_subtit"> 
                    <div align="center">Chamada</div></td>
                  <td width="450" class="tb_subtit"> 
                    <div align="center">Nome</div></td>
                  <td width="170" class="tb_subtit"> 
                    <div align="center">Data Nascimento</div></td>
                  <td width="60" class="tb_subtit"> 
                    <div align="center">Idade</div></td>
                  <td width="170" class="tb_subtit"> 
                    <div align="center">Situa&ccedil;&atilde;o 
                      da matr&iacute;cula</div></td>
                </tr>
                <%
				
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Matriculas where NU_Ano="& ano_letivo &" AND NU_Unidade="& unidade &" AND CO_Curso='"& curso &"' AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"' ORDER BY NU_Chamada"
		RS5.Open SQL5, CON1
chamadacheck=0
check=2
while not RS5.EOF
chamada= RS5("NU_Chamada")
matricula= RS5("CO_Matricula")
rematricula= RS5("DA_Rematricula")
situacao= RS5("CO_Situacao")
 if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if
  
if chamadacheck=chamada-1 then
%>
                <tr> 
                  <td width="10">
<div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  <td width="70" class="<%=cor%>"> 
                    <div align="center"><font class="form_dado_texto">  
                      <%

		response.Write(matricula)
%>
                      </font></div></td>
                  <td width="70" class="<%=cor%>"> 
                    <div align="center"><font class="form_dado_texto">  
                      <%
		response.Write(chamada)
%>
                      </font></div></td>
                  <td width="450" class="<%=cor%>"> 
                    <div align="center"><font class="form_dado_texto">  
                      <%

		Set RS8 = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT * FROM TB_Alunos where CO_Matricula="& matricula 
		RS8.Open SQL8, CON1

		Set RSCONTA = Server.CreateObject("ADODB.Recordset")
		SQLA = "SELECT * FROM TB_Contatos WHERE TP_Contato ='ALUNO' And CO_Matricula ="& matricula
		RSCONTA.Open SQLA, CONCONT

		no_al= RS8("NO_Aluno")	
nascimento = RSCONTA("DA_Nascimento_Contato")
if isnull(nascimento) then
dia_a = DatePart("d", now) 
mes_a = DatePart("m", now) 
ano_a = DatePart("yyyy", now)
else
vetor_nascimento = Split(nascimento,"/")  
dia_n = vetor_nascimento(0)
mes_n = vetor_nascimento(1)
ano_n = vetor_nascimento(2)

if dia_n<10 then 
dia_n = "0"&dia_n
end if

if mes_n<10 then
mes_n = "0"&mes_n
end if
dia_a = dia_n
mes_a = mes_n
ano_a = ano_n

end if
%>
                      <a href="../alunos/altera.asp?or=01&vd=vt&cod=<% =matricula %>&o=<%=obr%>" class='linkum'> 
                      <%		response.Write(no_al)%>
                      </a> </font> </div></td>
                  <td width="170" class="<%=cor%>"> 
                    <div align="center"><font class="form_dado_texto">  
                      <%

		response.Write(nascimento)
%>
                      </font></div></td>
                  <td width="60" class="<%=cor%>"> 
                    <div align="center"><font class="form_dado_texto">  
                      <%

					call aniversario(ano_a,mes_a,dia_a) 
					
					%>
                      </font></div></td>
                  <td width="170" class="<%=cor%>"> 
                    <div align="center"><font class="form_dado_texto">  
                      <%
		Set RS9 = Server.CreateObject("ADODB.Recordset")
		SQL9 = "SELECT * FROM TB_Situacao_Aluno where CO_Situacao='"& situacao&"'"
		RS9.Open SQL9, CON0
		
	no_situacao=RS9("TX_Descricao_Situacao")	
		
		response.Write(no_situacao)
%>
                      </font></div></td>
                </tr>
                <%
chamadacheck=chamadacheck+1
check=check+1
RS5.MOVENEXT
else
%>
                <tr> 
                  <td width="10" bgcolor="#CCCCCC"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                  <td width="70" bgcolor="#CCCCCC"></td>
                  <td width="70" bgcolor="#CCCCCC"></td>
                  <td width="450" bgcolor="#CCCCCC"></td>
                  <td width="170" bgcolor="#CCCCCC"></td>
                  <td width="60" bgcolor="#CCCCCC"></td>
                  <td width="170" bgcolor="#CCCCCC"></td>
                </tr>
                <tr> 
                  <td width="10">
<div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  <td width="70" class="<%=cor%>"> 
                    <div align="center"><font class="form_dado_texto">  
                      <%

		response.Write(matricula)
%>
                      </font></div></td>
                  <td width="70" class="<%=cor%>"> 
                    <div align="center"><font class="form_dado_texto">  
                      <%
		response.Write(chamada)
%>
                      </font></div></td>
                  <td width="450" class="<%=cor%>"> 
                    <div align="center"><font class="form_dado_texto">  
                      <%

		Set RS8 = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT * FROM TB_Alunos where CO_Matricula="& matricula 
		RS8.Open SQL8, CON1

		no_al= RS8("NO_Aluno")	
nascimento = RS8("DA_Nascimento")

vetor_nascimento = Split(nascimento,"/")  
dia_n = vetor_nascimento(0)
mes_n = vetor_nascimento(1)
ano_n = vetor_nascimento(2)

if dia_n<10 then 
dia_n = "0"&dia_n
end if

if mes_n<10 then
mes_n = "0"&mes_n
end if
dia_a = dia_n
mes_a = mes_n
ano_a = ano_n


%>
                      <a href="../alunos/altera.asp?or=01&vd=vt&cod=<% =matricula %>&o=<%=obr%>" class='linkum'> 
                      <%		response.Write(no_al)%>
                      </a> </font> </div></td>
                  <td width="170" class="<%=cor%>"> 
                    <div align="center"><font class="form_dado_texto">  
                      <%

		response.Write(nascimento)
%>
                      </font></div></td>
                  <td width="60" class="<%=cor%>"> 
                    <div align="center"><font class="form_dado_texto">  
                      <%

					call aniversario(ano_a,mes_a,dia_a) 
					
					%>
                      </font></div></td>
                  <td width="170" class="<%=cor%>"> 
                    <div align="center"><font class="form_dado_texto">  
                      <%
		Set RS9 = Server.CreateObject("ADODB.Recordset")
		SQL9 = "SELECT * FROM TB_Situacao_Aluno where CO_Situacao='"& situacao&"'"
		RS9.Open SQL9, CON0
		
	no_situacao=RS9("TX_Descricao_Situacao")	
		
		response.Write(no_situacao)
%>
                      </font></div></td>
                </tr>
                <%
chamadacheck=chamadacheck+2
check=check+2
RS5.MOVENEXT
end if

WEND
%>
                <tr> 
                  <td colspan="7"><div align="center"> 
                      <input name="bt" type="button" class="borda_bot3" id="bt" onClick="MM_goToURL('parent','carometro.asp?or=01&opt=vt&o=<%=obr%>');return document.MM_returnValue" value="Car&ocirc;metro">
                    </div></td>
                </tr>
              </table></td>
                  </tr>
                </table>
              </form></td>
  </tr>
  <tr>
    <td height="40"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

</body>
<script type="text/javascript">
<!--
  initInputHighlightScript();
//-->
</script>
<%end if 
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