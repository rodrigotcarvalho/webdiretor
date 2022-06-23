<%On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->

<!--#include file="../../../../inc/caminhos.asp"-->




<!--#include file="../../../../inc/funcoes2.asp"-->

<%
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
ma = request.form("ma")

unidade = request.Form("unidade")
obr=unidade&"_"&ma
		
		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON6 = Server.CreateObject("ADODB.Connection") 
		ABRIR6 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON6.Open ABRIR6
		
		Set CON7 = Server.CreateObject("ADODB.Connection") 
		ABRIR7 = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON7.Open ABRIR7				

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
//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900" background="../../../../img/fundo.gif" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)%>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr>                    
            
    <td height="10" class="tb_caminho"> <font class="style-caminho">
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
	  	call mensagens(nivel,636,0,0) 
	  end if%>
    </td>
                  </tr>				  				  

  <tr> 
    <td valign="top">
<table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
        <tr class="tb_tit"
> 
          <td width="653" height="15" class="tb_tit"
><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
            </strong></font></td>
        </tr>
        <tr> 
          <td><table width="1000" border="0" cellspacing="0">
              <tr> 
                <td width="10"><font class="form_dado_texto"> &nbsp;</font></td>
                <td width="495" class="tb_subtit"> 
                  <div align="center">UNIDADE </div></td>
                <td width="495" class="tb_subtit"> 
                  <div align="center">M&Ecirc;S </div></td>
              </tr>
              <tr> 
                <td width="10"> </td>
                <td width="495"> 
                  <div align="center"> <font class="form_dado_texto">  
                    <%
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")
					
					response.Write(no_unidade)%>
                    </font></div></td>
                <td width="495"> 
                  <div align="center"> <font class="form_dado_texto">  
                    <%
select case ma
 case 1 
 mes_a = "janeiro"
 case 2 
 mes_a = "fevereiro"
 case 3 
 mes_a = "março"
 case 4
 mes_a = "abril"
 case 5
 mes_a = "maio"
 case 6 
 mes_a = "junho"
 case 7
 mes_a = "julho"
 case 8 
 mes_a = "agosto"
 case 9 
 mes_a = "setembro"
 case 10 
 mes_a = "outubro"
 case 11 
 mes_a = "novembro"
 case 12 
 mes_a = "dezembro"
end select					  
response.Write(mes_a)%>
                    </font></div></td>
              </tr>
              <tr> 
                <td colspan="3">&nbsp;</td>
              </tr>
              <tr> 
                <td width="10"></td>
                <td colspan="2"><table width="1000" border="0" cellspacing="0">
                    <tr class="tb_subtit"> 
                      <td width="100">
<div align="center">Anivers&aacute;rio</div></td>
                      <td width="300"> 
                        <div align="center">Nome</div></td>
                      <td width="100">
<div align="center">Far&aacute;</div></td>
                      <td width="100">
<div align="center">Matr&iacute;cula</div></td>
                      <td width="200">
<div align="center">Curso</div></td>
                      <td width="100"> 
                        <div align="center">Etapa</div></td>
                      <td width="100"> 
                        <div align="center">Turma</div></td>
                    </tr>
                    <%

		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Contatos where TP_Contato='ALUNO' order by day(DA_Nascimento_Contato),Year(DA_Nascimento_Contato)"
		RS7.Open SQL7, CON7

While not RS7.EOF
cod_al=RS7("CO_Matricula")	  

		Set RSa0 = Server.CreateObject("ADODB.Recordset")
		SQLa0 = "SELECT * FROM TB_Matriculas where NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod_al &" AND NU_Unidade="&unidade
		RSa0.Open SQLa0, CON6
		
		Set RSa = Server.CreateObject("ADODB.Recordset")
		SQLa = "SELECT * FROM TB_Alunos where CO_Matricula ="& cod_al
		RSa.Open SQLa, CON6

IF RSa0.EOF	Then

RS7.Movenext
Else
nascimento = RS7("DA_Nascimento_Contato")		

if nascimento="" or isnull(nascimento) then

RS7.Movenext
else
nasceu=split(nascimento,"/")
dia_n=nasceu(0)

if nasceu(1)= ma then
nome_al=RSa("NO_Aluno")


		Set RSa1 = Server.CreateObject("ADODB.Recordset")
		SQLa1 = "SELECT * FROM TB_Matriculas where NU_Ano="& ano_letivo &" AND CO_Matricula="&cod_al
		RSa1.Open SQLa1, CON6
		
curso=RSa1("CO_Curso")
etp =RSa1("CO_Etapa")
tm =RSa1("CO_Turma")	
		
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
			
no_curso = RS1("NO_Curso")
		
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Etapa where CO_Curso ='"& curso &"' AND CO_Etapa ='"& etp &"'"
		RS2.Open SQL2, CON0

if RS2.eof	then
no_etp = "ETAPA SEM NOME CADASTRADO"
else
no_etp = RS2("NO_Etapa")
end if

 if dia_n mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if
%>
                    <tr> 
                      <td width="100" class="<%=cor%>">
<div align="center"><font class="form_dado_texto">  
                          <%response.Write(nasceu(0))%>
                          </font></div></td>
                      <td width="300" class="<%=cor%>"> 
                        <div align="center"><font class="form_dado_texto">  
                          <%response.Write(nome_al)%>
                          </font></div></td>
                      <td width="100" class="<%=cor%>">
<div align="center"><font class="form_dado_texto">  
                          <%call aniversario(nasceu(2),nasceu(1),nasceu(0)) %>
                          </font></div></td>
                      <td width="100" class="<%=cor%>"> 
                        <div align="center"><font class="form_dado_texto">  
                          <%response.Write(cod_al)%>
                          </font></div></td>
                      <td width="200" class="<%=cor%>">
<div align="center"><font class="form_dado_texto">  
                          <%response.Write(no_curso)%>
                          </font></div></td>
                      <td width="100" class="<%=cor%>"> 
                        <div align="center"><font class="form_dado_texto">  
                          <%response.Write(no_etp)%>
                          </font></div></td>
                      <td width="100" class="<%=cor%>"> 
                        <div align="center"><font class="form_dado_texto"> 
                          <%response.Write(tm)%>
                          </font></div></td>
                    </tr>
                    <%

RS7.Movenext
Else

RS7.Movenext
end if
end if
end if
Wend
%>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
<%call GravaLog (chave,obr)%>
</body>
<%'end if %>
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