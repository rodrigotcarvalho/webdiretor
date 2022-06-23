<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->

<% 
opt= request.QueryString("opt")
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
ori = request.QueryString("or")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
cod= request.QueryString("cod_cons")	

z = request.QueryString("z")
erro = request.QueryString("e")
vindo = request.QueryString("vd")
obr = request.QueryString("o")

if vindo="crmt" then
dados= split(obr, "_" )
unidade= dados(0)
curso= dados(1)
co_etapa= dados(2)
turma= dados(3)
end if
obr=cod
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
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod
		RS.Open SQL, CON1
		
		
codigo = RS("CO_Matricula")
nome_prof = RS("NO_Aluno")



		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod
		RS.Open SQL, CON1


ano_aluno = RS("NU_Ano")
rematricula = RS("DA_Rematricula")
situacao = RS("CO_Situacao")
encerramento= RS("DA_Encerramento")
unidade= RS("NU_Unidade")
curso= RS("CO_Curso")
etapa= RS("CO_Etapa")
turma= RS("CO_Turma")
cham= RS("NU_Chamada")




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
  if (document.busca.busca1.value != "" && document.busca.busca2.value != "")
  {    alert("Por favor digite SOMENTE uma opção de busca!")
    document.busca.busca1.focus()
    return false
  }
    if (document.busca.busca1.value == "" && document.busca.busca2.value == "")
  {    alert("Por favor digite uma opção de busca!")
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
//-->
</script>
</head>
<% if opt="listall" or opt="list" then%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%else %>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%end if %>
<%call cabecalho(nivel)
%>
<div id="fundo" style="position:absolute; left:0px; top:0px; width:100%; height:100%; z-index:1; background-color: #000000; layer-background-color: #000000; border: 1px none #000000; visibility: hidden;" class="transparente"></div>
<div id="alinha" style="position:absolute; width:400px; visibility: hidden; z-index: 2; left: 326px; height: 520px;"> 
  <table width="300" border="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr> 
      <td width="478"> <div align="right"> <span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')">fechar</a></font></span></div></td>
      <td width="20"><div align="right"><span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')"><img src="../../../../img/fecha.gif" width="20" height="16" border="0"></a></font></span></div></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center" ><img src="../../../../img/fotos/aluno/<% =codigo %>.jpg" height="500"></div></td>
    </tr>
    <tr>
      <td colspan="2"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
          <%response.Write(nome_prof)%>
          </font></div></td>
    </tr>
  </table>
</div>

<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
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
      <%call mensagens(nivel,636,0,0) %>
    </td>
			  </tr>			  
        <form action="cadastro.asp?opt=list&or=01" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
          <tr>
      <td valign="top"> 
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
          <tr> 
            <td width="653" class="tb_tit"
>Dados Escolares</td>
            <td width="113" class="tb_tit"
> </td>
          </tr>
          <tr> 
            <td height="10"> <table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="19%" height="10"> <div align="right"><font class="form_dado_texto"> 
                      Matr&iacute;cula: </font></div></td>
                  <td width="9%" height="10"><font class="form_dado_texto"> 
                    <input name="cod" type="hidden" value="<%=codigo%>">
                    <%response.Write(codigo)%>
                    </font></td>
                  <td width="6%" height="10"> <div align="right"><font class="form_dado_texto"> 
                      Nome: </font></div></td>
                  <td width="66%" height="10"><font class="form_dado_texto"> 
                    <%response.Write(nome_prof)%>
                    <input name="nome2" type="hidden" class="textInput" id="nome2"  value="<%response.Write(nome_prof)%>" size="75" maxlength="50">
                    &nbsp;</font></td>
                </tr>
              </table></td>
            <td valign="top">&nbsp; </td>
          </tr>
          <tr> 
            <td height="10" bgcolor="#FFFFFF">&nbsp;</td>
            <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="2"><table width="100%" border="0" cellspacing="0">
                <tr class="tb_subtit"> 
                  <td width="33" height="10"> <div align="center"> 
                      <%
call GeraNomes("PORT",unidade,curso,etapa,CON0)
no_unidades = session("no_unidades")
no_grau = session("no_grau")
no_serie = session("no_serie")
%>
                      Ano</div></td>
                  <td width="81" height="10"> <div align="center">Matr&iacute;cula</div></td>
                  <td width="75" height="10"> <div align="center">Cancelamento</div></td>
                  <td width="86" height="10"> <div align="center"> Situa&ccedil;&atilde;o</div></td>
                  <td width="113" height="10"> <div align="center">Unidade</div></td>
                  <td width="133" height="10"> <div align="center">Curso</div></td>
                  <td width="85" height="10"> <div align="center"> Etapa</div></td>
                  <td width="90" height="10"> <div align="center">Turma </div></td>
                  <td width="54" height="10"> <div align="center">Chamada</div></td>
                </tr>
                <tr class="tb_corpo"
> 
                  <td width="33" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(ano_aluno)%>
                      </font></div></td>
                  <td width="81" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(rematricula)%>
                      </font></div></td>
                  <td width="75" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(encerramento)%>
                      </font></div></td>
                  <td width="86" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%
					
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
		RSCONTST.Open SQLCONTST, CON0
							
				no_situacao = RSCONTST("TX_Descricao_Situacao")	
					response.Write(no_situacao)%>
                      </font></div></td>
                  <td width="113" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_unidades)%>
                      </font></div></td>
                  <td width="133" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_grau)%>
                      </font></div></td>
                  <td width="85" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_serie)%>
                      </font></div></td>
                  <td width="90" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(turma)%>
                      </font></div></td>
                  <td width="54" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(cham)%>
                      </font></div></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td height="10" colspan="2" class="tb_tit"
>Avalia&ccedil;&otilde;es</td>
          </tr>
          <tr> 
            <td colspan="2"><table width="100%" border="0" align="center" cellspacing="0" class="tb_corpo"
>
                <tr> 
                  <td> 
                    <%		Set RS_tb = Server.CreateObject("ADODB.Recordset")
		SQL_tb = "SELECT * FROM TB_Da_Aula WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso&"' AND CO_Etapa ='"& etapa&"' AND CO_Turma ='"& turma&"'"
		RS_tb.Open SQL_tb, CON2

if RS_tb.eof then
%>
                    <div align="center"> <font class="style1"> 
                      <%response.Write("<br><br><br><br><br>Não existe Boletim para este aluno!")%>
                      </font></div>
                    <%
else
notaFIL=RS_tb("TP_Nota")



if notaFIL ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na

elseif notaFIL="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb

elseif notaFIL ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc
else
		response.Write("ERRO")
end if	


if tb_nota="TB_NOTA_A" then
minimo_recuperacao= 60
end if		
%>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="180" rowspan="2" class="tb_subtit"
> <div align="left"><strong>Disciplina</strong></div></td>
                        <td colspan="5" class="tb_subtit"
> <div align="center">TRIMESTRE 1</div></td>
                        <td colspan="5" class="tb_subtit"
><div align="center">TRIMESTRE 2</div></td>
                        <td colspan="5" class="tb_subtit"
><div align="center">TRIMESTRE 3</div></td>
                        <td width="50" rowspan="2" class="tb_subtit"
> <div align="center">RA </div>
                          <div align="center"></div></td>
                        <td colspan="4" class="tb_subtit"
><div align="center">ETAPA COMPLEMENTAR</div></td>
                        <td width="50" rowspan="2" class="tb_subtit"
> <div align="center">RF</div></td>
                      </tr>
                      <tr> 
                        <td width="36" class="tb_subtit"
> <div align="center">SAPR</div></td>
                        <td width="36" class="tb_subtit"
> <div align="center">PR</div></td>
                        <td width="36" class="tb_subtit"
> <div align="center">MP</div></td>
                        <td width="36" class="tb_subtit"
> <div align="center">MC</div></td>
                        <td width="36" class="tb_subtit"
> <div align="center">F</div></td>
                        <td width="36" class="tb_subtit"
> <div align="center">SAPR</div></td>
                        <td width="36" class="tb_subtit"
> <div align="center">PR</div></td>
                        <td width="36" class="tb_subtit"
> <div align="center">MP</div></td>
                        <td width="36" class="tb_subtit"
> <div align="center">MC*</div></td>
                        <td width="36" class="tb_subtit"
> <div align="center">F</div></td>
                        <td width="36" class="tb_subtit"
> <div align="center">SAPR</div></td>
                        <td width="36" class="tb_subtit"
> <div align="center">PR</div></td>
                        <td width="36" class="tb_subtit"
> <div align="center">MP</div></td>
                        <td width="36" class="tb_subtit"
> <div align="center">MC</div></td>
                        <td width="36" class="tb_subtit"
> <div align="center">F</div></td>
                        <td width="45" class="tb_subtit"
> <div align="center">SAPR</div></td>
                        <td width="45" class="tb_subtit"
> <div align="center">PR</div></td>
                        <td width="45" class="tb_subtit"
> <div align="center">MP</div></td>
                        <td width="45" class="tb_subtit"
> <div align="center">MC</div></td>
                      </tr>
                      <%
rec_lancado="sim"

		Set RSprog = Server.CreateObject("ADODB.Recordset")
		SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
		RSprog.Open SQLprog, CON0

check=2
	
while not RSprog.EOF

	materia=RSprog("CO_Materia")
	mae=RSprog("IN_MAE")
	fil=RSprog("IN_FIL")
	in_co=RSprog("IN_CO")
	nu_peso=RSprog("NU_Peso")
	ordem=RSprog("NU_Ordem_Boletim")

		Set RS1a = Server.CreateObject("ADODB.Recordset")
		SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
		RS1a.Open SQL1a, CON0
		
no_materia=RS1a("NO_Materia")

 if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if

		
		
		Set CON_N = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIRn

'for periodofil=1 to 4


		
		Set RSnFIL = Server.CreateObject("ADODB.Recordset")
	  	Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"'"
		Set RS3 = CON_N.Execute(SQL_N)



if RS3.EOF then
va_apr1="&nbsp;"
va_apr2="&nbsp;"
va_apr3="&nbsp;"
va_apr4="&nbsp;"
va_apr5="&nbsp;"
va_apr6="&nbsp;"
va_apr7="&nbsp;"
va_apr8="&nbsp;"
va_v_apr1="&nbsp;"
va_v_apr2="&nbsp;"
va_v_apr3="&nbsp;"
va_v_apr4="&nbsp;"
va_v_apr5="&nbsp;"
va_v_apr6="&nbsp;"
va_v_apr7="&nbsp;"
va_v_apr8="&nbsp;"
va_sapr="&nbsp;"
va_pr="&nbsp;"
va_te="&nbsp;"
va_bon="&nbsp;"
va_me="&nbsp;"
va_mc="&nbsp;"
va_faltas="&nbsp;"	
data_grav="nulo"
hora_grav="nulo"	
va_sapr1="&nbsp;"
va_pr1="&nbsp;"
va_te1="&nbsp;"
pr1="&nbsp;"
va_me1="&nbsp;"
va_mc1="&nbsp;"
va_faltas1="&nbsp;"
va_sapr2="&nbsp;"
va_pr2="&nbsp;"
va_te2="&nbsp;"
pr2="&nbsp;"
va_me2="&nbsp;"
va_mc2="&nbsp;"
va_faltas2="&nbsp;"
va_sapr3="&nbsp;"
va_pr3="&nbsp;"
va_te3="&nbsp;"
pr3="&nbsp;"
va_me3="&nbsp;"
va_mc3="&nbsp;"
va_faltas3="&nbsp;"
va_sapr4="&nbsp;"
va_pr4="&nbsp;"
va_me4="&nbsp;"
va_mc4="&nbsp;"
	
	
else
va_sapr1=RS3("VA_Sapr1")
va_pr1=RS3("VA_Pr1")
va_te1=RS3("VA_Te1")
va_me1=RS3("VA_Me1")
va_mc1=RS3("VA_Mc1")
va_faltas1=RS3("NU_Faltas_P1")
va_pr1=va_pr1*1
va_te1=va_te1*1
pr1=va_pr1+va_te1

va_sapr2=RS3("VA_Sapr2")
va_pr2=RS3("VA_Pr2")
va_te2=RS3("VA_Te2")
va_me2=RS3("VA_Me2")
va_mc2=RS3("VA_Mc2")
va_faltas2=RS3("NU_Faltas_P2")

va_pr2=va_pr2*1
va_te2=va_te2*1
pr2=va_pr2+va_te2

va_sapr3=RS3("VA_Sapr3")
va_pr3=RS3("VA_Pr3")
va_te3=RS3("VA_Te3")
va_me3=RS3("VA_Me3")
va_mc3=RS3("VA_Mc3")
va_faltas3=RS3("NU_Faltas_P3")

va_pr3=va_pr3*1
va_te3=va_te3*1
pr3=va_pr3+va_te3

va_sapr4=RS3("VA_Sapr_EC")
va_pr4=RS3("VA_Pr4")
va_me4=RS3("VA_Me_EC")
va_mc4=RS3("VA_Mfinal")

pr4=va_pr

		
'		Set RS4 = Server.CreateObject("ADODB.Recordset")
'		SQL4 = "SELECT * FROM TB_Controle"
'		RS4.Open SQL4, CON
	
'co_apr1=RS4("CO_apr1")
'co_apr2=RS4("CO_apr2")
'co_apr3=RS4("CO_apr3")
'co_apr4=RS4("CO_apr4")
'co_prova1=RS4("CO_prova1")
'co_prova2=RS4("CO_prova2")
'co_prova3=RS4("CO_prova3")
'co_prova4=RS4("CO_prova4")	
		
'if periodo_check=1 then		
'		if co_apr1="D"then
'		showapr1="n"
'		else 
		showapr1="s"
'		end if
'		if co_prova1="D"then
'		showprova1="n"
'		else 
		showprova1="s"
'		end if
'elseif periodo_check=2 then		
'		if co_apr2="D"then
'		showapr2="n"
'		else 
		showapr2="s"
'		end if
'		if co_prova2="D"then
'		showprova2="n"
'		else 
		showprova2="s"
'		end if					
'elseif periodo_check=3 then		
'		if co_apr3="D"then
'		showapr3="n"
'		else 
		showapr3="s"
'		end if
'		if co_prova3="D"then
'		showprova3="n"
'		else 
		showprova3="s"
'		end if
'elseif periodo_check=4 then		
'		if co_apr4="D"then
'		showapr4="n"
'		else 
		showapr4="s"
'		end if
'		if co_prova4="D"then
'		showprova4="n"
'		else 
		showprova4="s"
'		end if
'end if											
		
		
				
end if


if va_me1="" or va_me1="&nbsp;" or isnull(va_me1)then
else
va_me1=va_me1/10
'	decimo = va_me1 - Int(va_me1)
'		If decimo >= 0.5 Then
'			nota_arredondada = Int(va_me1) + 1
'			va_me1=nota_arredondada
'		Else
'			nota_arredondada = Int(va_me1)
'			va_me1=nota_arredondada					
'		End If
	va_me1 = formatNumber(va_me1,1)
end if	
	
if va_mc1="" or va_mc1="&nbsp;" or isnull(va_mc1)then
else	
va_mc1=va_mc1/10
'	decimo = va_mc1 - Int(va_mc1)
'		If decimo >= 0.5 Then
'			nota_arredondada = Int(va_mc1) + 1
'			va_mc1=nota_arredondada
'		Else
'			nota_arredondada = Int(va_mc1)
'			va_mc1=nota_arredondada					
'		End If
	va_mc1 = formatNumber(va_mc1,1)
end if	
if va_me2=" "or va_me2="&nbsp;" or isnull(va_me2)then
else
va_me2=va_me2/10
'	decimo = va_me2 - Int(va_me2)
'		If decimo >= 0.5 Then
'			nota_arredondada = Int(va_me2) + 1
'			va_me2=nota_arredondada
'		Else
'			nota_arredondada = Int(va_me2)
'			va_me2=nota_arredondada					
'		End If
	va_me2 = formatNumber(va_me2,1)
end if	
	
if va_mc2="" or va_mc2="&nbsp;" or isnull(va_mc2)then
else		
va_mc2=va_mc2/10
'	decimo = va_mc2 - Int(va_mc2)
'		If decimo >= 0.5 Then
'			nota_arredondada = Int(va_mc2) + 1
'			va_mc2=nota_arredondada
'		Else
'			nota_arredondada = Int(va_mc2)
'			va_mc2=nota_arredondada					
'		End If
	va_mc2 = formatNumber(va_mc2,1)		
end if	
if va_me3="" or va_me3="&nbsp;" or isnull(va_me3)then
else
va_me3=va_me3/10
'	decimo = va_me3 - Int(va_me3)
'		If decimo >= 0.5 Then
'			nota_arredondada = Int(va_me3) + 1
'			va_me3=nota_arredondada
'		Else
'			nota_arredondada = Int(va_me3)
'			va_me3=nota_arredondada					
'		End If
	va_me3 = formatNumber(va_me3,1)
end if	
	
if va_mc3="" or va_mc3="&nbsp;" or isnull(va_mc3)then
else		

va_mc3=va_mc3/10
'	decimo = va_mc3 - Int(va_mc3)
'		If decimo >= 0.5 Then
'			nota_arredondada = Int(va_mc3) + 1
'			va_mc3=nota_arredondada
'		Else
'			nota_arredondada = Int(va_mc3)
'			va_mc3=nota_arredondada					
'		End If
	va_mc3 = formatNumber(va_mc3,1)
end if	
if va_me4="" or va_me4="&nbsp;" or isnull(va_me4)then
else	
va_me4=va_me4/10
'	decimo = va_me4 - Int(va_me4)
'		If decimo >= 0.5 Then
'			nota_arredondada = Int(va_me4) + 1
'			va_me4=nota_arredondada
'		Else
'			nota_arredondada = Int(va_me4)
'			va_me4=nota_arredondada					
'		End If
	va_me4 = formatNumber(va_me4,1)
end if	
	
if va_mc4="" or va_mc4="&nbsp;" or isnull(va_mc4)then
else		
va_mc4=va_mc4/10
'	decimo = va_mc4 - Int(va_mc4)
'		If decimo >= 0.5 Then
'			nota_arredondada = Int(va_mc4) + 1
'			va_mc4=nota_arredondada
'		Else
'			nota_arredondada = Int(va_mc4)
'			va_mc4=nota_arredondada					
'		End If
	va_mc4 = formatNumber(va_mc4,1)	
end if

if (isnull(va_apr1) OR va_apr1="&nbsp;") and (ISNULL(va_apr2) OR va_apr2="&nbsp;")and (ISNULL(va_apr3) OR va_apr3="&nbsp;")and (ISNULL(va_apr4) OR va_apr4="&nbsp;")and (ISNULL(va_apr5)  OR va_apr5="&nbsp;")and (ISNULL(va_apr6) OR  va_apr6="&nbsp;") and (ISNULL(va_apr7) OR va_apr7="&nbsp;")and (ISNULL(va_apr8) OR va_apr8="&nbsp;")and (ISNULL(va_sapr) OR va_sapr="&nbsp;")  then
data_inicio=""
va_faltas=""
else
		if (va_apr1=0 OR va_apr1="0") and (va_apr2=0 OR va_apr2="0")and (va_apr3=0 OR va_apr3="0")and (va_apr4=0 OR va_apr4="0")and (va_apr5=0 OR va_apr5="0")and (va_apr6=0 OR va_apr6="0") and (va_apr7=0 OR va_apr7="0") and (va_apr8=0 OR va_apr8="0")and (va_sapr=0 OR va_sapr="0")  then
		data_inicio=""
		va_faltas=""
		end if
end if




%>
                      <tr class="<%response.Write(cor)%>"> 
                        <td width="265" class="<%response.Write(cor)%>" 
> 
            <%response.Write(no_materia)%>
          </td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
							if showapr1="s" and showprova1="s" then
	if va_sapr1="&nbsp;" or isnull(va_sapr1) then
	else						
	va_sapr1 = formatNumber(va_sapr1,1)
	end if												
							response.Write("&nbsp;"&va_sapr1)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
	if showapr1="s" and showprova1="s" then
		if pr1="&nbsp;" or isnull(pr1) OR pr1="" then
			response.Write("&nbsp;")
		else							
			pr1 = formatNumber(pr1,1)												
			response.Write("&nbsp;"&pr1)
		end if							
	else
	response.Write("&nbsp;")							
	end if
							%>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
							if showapr1="s" and showprova1="s" then					
							response.Write("&nbsp;"&va_me1)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
							if showapr1="s" and showprova1="s" then					
							response.Write("&nbsp;"&va_mc1)
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
							if showapr1="s" and showprova1="s" then
							response.Write("&nbsp;"&va_faltas1)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
							if showapr2="s" and showprova2="s" then
	if va_sapr2="&nbsp;" or isnull(va_sapr2) then
								response.Write("&nbsp;")
	else							
	va_sapr2 = formatNumber(va_sapr2,1)												
							response.Write("&nbsp;"&va_sapr2)
	end if							
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
							if showapr2="s" and showprova2="s" then
	if pr2="&nbsp;" or isnull(pr2) then
								response.Write("&nbsp;")
	else							
	pr2 = formatNumber(pr2,1)												
							response.Write("&nbsp;"&pr2)
	end if							
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
							if showapr2="s" and showprova2="s" then					
							response.Write("&nbsp;"&va_me2)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
							if showapr2="s" and showprova2="s" then					
							response.Write("&nbsp;"&va_mc2)
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
							if showapr2="s" and showprova2="s" then
							response.Write("&nbsp;"&va_faltas2)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
							if showapr3="s" and showprova3="s" then
	if va_sapr3="&nbsp;" or isnull(va_sapr3) then
								response.Write("&nbsp;")
	else							
	va_sapr3 = formatNumber(va_sapr3,1)												
							response.Write("&nbsp;"&va_sapr3)
	end if		
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
							if showapr3="s" and showprova3="s" then
	if pr3="&nbsp;" or isnull(pr3) then
								response.Write("&nbsp;")
	else							
	pr3 = formatNumber(pr3,1)												
							response.Write("&nbsp;"&pr3)
	end if								

							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
							if showapr3="s" and showprova3="s" then					
							response.Write("&nbsp;"&va_me3)
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
							if showapr3="s" and showprova3="s" then					
							response.Write("&nbsp;"&va_mc3)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
							if showapr3="s" and showprova3="s" then
							response.Write("&nbsp;"&va_faltas3)
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
					if showapr3="s" and showprova3="s" then	
						if va_mc3="&nbsp;" or isnull(va_mc3) then
							response.Write(va_mc3)
						else
							if va_mc3 < 7 then					
							response.Write("ECE")
							resultado1="ece"
							else
							response.Write("APR")
							resultado1="apr"							
							end if
						end if	
					else
					response.Write("&nbsp;")							
					end if							
							%>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
			  if resultado1="ece" then 			  
							if showapr4="s" and showprova4="s" then
	if va_sapr4="&nbsp;" or isnull(va_sapr4) then
								response.Write("&nbsp;")
	else							
	va_sapr4 = formatNumber(va_sapr4,1)
							response.Write("&nbsp;"&va_sapr4)
	end if													

							else
							response.Write("&nbsp;")
							end if
			else
				response.Write("&nbsp;")
			end if							
%>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
			  if resultado1="ece" then 
				if showapr4="s" and showprova4="s" then
					if pr4="&nbsp;" or isnull(pr4) OR pr4="" then
						response.Write("&nbsp;")
					else							
						pr4 = formatNumber(pr4,1)								
							response.Write("&nbsp;"&pr4)
					end if								
				else
					response.Write("&nbsp;")
				end if
			else
				response.Write("&nbsp;")
			end if						  
			  %>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
			  if resultado1="ece" then 			  
			if showapr4="s" and showprova4="s" then					
				response.Write("&nbsp;"&va_me4)			
			else			
				response.Write("&nbsp;")
			end if
			else
				response.Write("&nbsp;")
			end if			
			%>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
			  if resultado1="ece" then 			  
			if showapr4="s" and showprova4="s"then					
				response.Write("&nbsp;"&va_mc4)							
			else			  
			response.Write("&nbsp;")
			end if
			else
				response.Write("&nbsp;")
			end if			
			  %>
            </div></td>
          <td width="35" class="<%response.Write(cor)%>" 
> <div align="center"> 
              <%
			  if resultado1="ece" then 			  
					if showapr4="s" and showprova4="s" then	
						if va_mc4="&nbsp;" or isnull(va_mc4) then
							response.Write("&nbsp;")
						else
							if va_mc4 < 5 then					
							response.Write("REP")
							else
							response.Write("APR")
							end if
						end if	
					else							
					response.Write("&nbsp;")							
					end if
			else
				response.Write("&nbsp;")
			end if					
							%>
            </div></td>
                      </tr>
                      <%check=check+1
RSprog.MOVENEXT
wend%>
                      <tr valign="bottom"> 
                        <td height="20" colspan="22" 
> <div align="right"><font class="style1">Sapr–Média das Aprs, PR-Prova, MP-Média 
                            do Período, MC-Média Acumulada, F-Faltas, RA-Resultado 
                            Anual, RF-Resultado Final</font></div></td>
                      </tr>
                      <tr valign="bottom">
                        <td height="20" colspan="22" 
>  <div align="right"><font class="style1">* Esta nota est&aacute; sujeita a altera&ccedil;&otilde;es pela 1&ordf; 
                          Etapa Complementar de Estudos (Vide o Boletim de Avalia&ccedil;&otilde;es 
                          do 2&ordm; Trimestre).</font></div></td>
                      </tr>
                    </table>
                    <%end if%>
                  </td>
                </tr>
              </table></td>
          </tr>
          <tr class="tb_tit"
> 
            <td colspan="2">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
</form>
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

</body>

<%call GravaLog (chave,obr)%>
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