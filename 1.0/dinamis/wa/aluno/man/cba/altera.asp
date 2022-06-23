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
	

z = request.QueryString("z")
erro = request.QueryString("e")
vindo = request.QueryString("vd")
obr = request.QueryString("o")

if vindo="1" then
periodo_check=request.form("periodo")
cod= request.form("cod")
else
cod= request.QueryString("cod_cons")
periodo_check=1
end if
obr=cod&"?"&periodo_check


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
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,636,0,0) %>
    </td>
			  </tr>			  
        <form action="altera.asp?vd=1" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
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
                  <td width="19%" height="10"> <div align="right"><font class="form_dado_texto"> Matr&iacute;cula: 
                      </font></div></td>
                  <td width="9%" height="10"><font class="form_dado_texto"> 
                    <input name="cod" type="hidden" value="<%=codigo%>">
                    <%response.Write(codigo)%>
                   </font></td>
                  <td width="6%" height="10"> <div align="right"><font class="form_dado_texto"> Nome: 
                      </font></div></td>
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
                  <td width="40" height="10"> 
                    <div align="center"> 
                      <%
call GeraNomes("PORT",unidade,curso,etapa,CON0)
no_unidades = session("no_unidades")
no_grau = session("no_grau")
no_serie = session("no_serie")
%>
                      Ano</div></td>
                  <td width="80" height="10"> 
                    <div align="center">Matr&iacute;cula</div></td>
                  <td width="100" height="10"> 
                    <div align="center">Cancelamento</div></td>
                  <td width="100" height="10"> 
                    <div align="center"> Situa&ccedil;&atilde;o</div></td>
                  <td width="100" height="10"> 
                    <div align="center">Unidade</div></td>
                  <td width="130" height="10"> 
                    <div align="center">Curso</div></td>
                  <td width="100" height="10"> 
                    <div align="center"> Etapa</div></td>
                  <td width="100" height="10"> 
                    <div align="center">Turma </div></td>
                  <td width="100" height="10"> 
                    <div align="center">Chamada</div></td>
                  <td width="150"> 
                    <div align="center">Per&iacute;odo</div></td>
                </tr>
                <tr class="tb_corpo"
> 
                  <td width="40" height="10"> 
                    <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(ano_aluno)%>
                      </font></div></td>
                  <td width="80" height="10"> 
                    <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(rematricula)%>
                      </font></div></td>
                  <td width="100" height="10"> 
                    <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(encerramento)%>
                      </font></div></td>
                  <td width="100" height="10"> 
                    <div align="center"> <font class="form_dado_texto"> 
                      <%
					
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
		RSCONTST.Open SQLCONTST, CON0
							
				no_situacao = RSCONTST("TX_Descricao_Situacao")	
					response.Write(no_situacao)%>
                      </font></div></td>
                  <td width="100" height="10"> 
                    <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_unidades)%>
                      </font></div></td>
                  <td width="130" height="10"> 
                    <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_grau)%>
                      </font></div></td>
                  <td width="100" height="10"> 
                    <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_serie)%>
                      </font></div></td>
                  <td width="100" height="10"> 
                    <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(turma)%>
                      </font></div></td>
                  <td width="100" height="10"> 
                    <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(cham)%>
                      </font></div></td>
                  <td width="150"> 
                    <div align="center">
                      <select name="periodo" class="borda" id="periodo" onChange="MM_callJS('submitfuncao()')">
                        <%
		Set RSPER = Server.CreateObject("ADODB.Recordset")
		SQLPER = "SELECT * FROM TB_Periodo order by NU_Periodo"'"
		RSPER.Open SQLPER, CON0
		
		While not RSPER.EOF
		periodo=RSPER("NU_Periodo")
		no_periodo=RSPER("NO_Periodo")
		periodo=periodo*1
		periodo_check=periodo_check*1
		
		if periodo=periodo_check then		
		%>
                        <option value="<%=periodo%>" selected><%=no_periodo%></option>
		<%else%>
		                <option value="<%=periodo%>"><%=no_periodo%></option>				
		<%end if
		RSPER.Movenext
		WEND
		%>
                      </select>
                    </div></td>
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
            <td colspan="2">		

			  </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2"><div align="right">
                <table width="100%" border="0" align="center" cellspacing="0" class="tb_corpo"
>
                  <tr> 
                    <td> 
                      <%		Set RS_tb = Server.CreateObject("ADODB.Recordset")
		SQL_tb = "SELECT * FROM TB_Da_Aula WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso&"' AND CO_Etapa ='"& etapa&"' AND CO_Turma ='"& turma&"'"
		RS_tb.Open SQL_tb, CON2

if RS_tb.eof then
%>
                      <div align="center"> <font class="style1"> 
                        <%response.Write("<br><br><br><br><br>Não existe Avaliações Progressivas para este aluno!")%>
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

if periodo_check=2 then
width=85
else
width=115
end if
%>
                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td width="<%response.Write(width)%>" rowspan="2" class="tb_subtit"
> <div align="left"><strong>Disciplina</strong></div></td>
                          <td width="30" rowspan="2" class="tb_subtit"
> <div align="center">F</div></td>
                          <td colspan="2" class="tb_subtit"
><div align="center">APR1</div></td>
                          <td colspan="2" class="tb_subtit"
><div align="center">APR2</div></td>
                          <td colspan="2" class="tb_subtit"
><div align="center">APR3</div></td>
                          <td colspan="2" class="tb_subtit"
><div align="center">APR4</div></td>
                          <td colspan="2" class="tb_subtit"
><div align="center">APR5</div></td>
                          <td colspan="2" class="tb_subtit"
><div align="center">APR6</div></td>
                          <td colspan="2" class="tb_subtit"
><div align="center">TEC1</div></td>
                          <td colspan="2" class="tb_subtit"
><div align="center">TEC2</div></td>
                          <td width="30" rowspan="2" class="tb_subtit"
> <div align="center">SAPR</div></td>
                          <td width="30" rowspan="2" class="tb_subtit"
> <div align="center">PR</div></td>
                          <td width="30" rowspan="2" class="tb_subtit"
> <div align="center">MP</div></td>
<% if periodo_check=2 then%>
                          <td width="30" rowspan="2" class="tb_subtit"
> <div align="center">EC1</div></td>
<%end if%>
                          <td width="170" rowspan="2" class="tb_subtit"
><div align="center">Alterado por</div></td>
                          <td width="115" rowspan="2" class="tb_subtit"
> <div align="center">Data/Hora</div></td>
                        </tr>
                        <tr> 
                          <td width="30" class="tb_subtit"
> <div align="center">N</div></td>
                          <td width="30" class="tb_subtit"
> <div align="center">P</div></td>
                          <td width="30" class="tb_subtit"
> <div align="center">N</div></td>
                          <td width="30" class="tb_subtit"
> <div align="center">P</div></td>
                          <td width="30" class="tb_subtit"
> <div align="center">N</div></td>
                          <td width="30" class="tb_subtit"
> <div align="center">P</div></td>
                          <td width="30" class="tb_subtit"
> <div align="center">N</div></td>
                          <td width="30" class="tb_subtit"
> <div align="center">P</div></td>
                          <td width="30" class="tb_subtit"
> <div align="center">N</div></td>
                          <td width="30" class="tb_subtit"
> <div align="center">P</div></td>
                          <td width="30" class="tb_subtit"
> <div align="center">N</div></td>
                          <td width="30" class="tb_subtit"
> <div align="center">P</div></td>
                          <td width="30" class="tb_subtit"
> <div align="center">N</div></td>
                          <td width="30" class="tb_subtit"
> <div align="center">P</div></td>
                          <td width="30" class="tb_subtit"
> <div align="center">N</div></td>
                          <td width="30" class="tb_subtit"
> <div align="center">P</div></td>
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
va_apr1=""
va_apr2=""
va_apr3=""
va_apr4=""
va_apr5=""
va_apr6=""
va_apr7=""
va_apr8=""
va_v_apr1=""
va_v_apr2=""
va_v_apr3=""
va_v_apr4=""
va_v_apr5=""
va_v_apr6=""
va_v_apr7=""
va_v_apr8=""
va_sapr=""
va_pr=""
va_te=""
va_bon=""
va_me=""
va_mc=""
va_faltas=""
usuario_grav=""	
data_grav="nulo"
hora_grav="nulo"	
else
if periodo_check=1 then
va_apr1=RS3("Apr1_P1")
va_apr2=RS3("Apr2_P1")
va_apr3=RS3("Apr3_P1")
va_apr4=RS3("Apr4_P1")
va_apr5=RS3("Apr5_P1")
va_apr6=RS3("Apr6_P1")
va_apr7=RS3("Apr7_P1")
va_apr8=RS3("Apr8_P1")
va_v_apr1=RS3("V_Apr1_P1")
va_v_apr2=RS3("V_Apr2_P1")
va_v_apr3=RS3("V_Apr3_P1")
va_v_apr4=RS3("V_Apr4_P1")
va_v_apr5=RS3("V_Apr5_P1")
va_v_apr6=RS3("V_Apr6_P1")
va_v_apr7=RS3("V_Apr7_P1")
va_v_apr8=RS3("V_Apr8_P1")
va_sapr=RS3("VA_Sapr1")
va_pr=RS3("VA_Pr1")
va_te=RS3("VA_Te1")
va_bon=RS3("VA_Bon1")
va_me=RS3("VA_Me1")
va_mc=RS3("VA_Mc1")
va_faltas=RS3("NU_Faltas_P1")
usuario_grav= RS3("CO_Usuario")
data_grav=RS3("DA_Ult_Acesso")
hora_grav=RS3("HO_ult_Acesso")

va_pr=va_pr*1
va_te=va_te*1
pr=va_pr+va_te
elseif periodo_check=2 then
va_apr1=RS3("Apr1_P2")
va_apr2=RS3("Apr2_P2")
va_apr3=RS3("Apr3_P2")
va_apr4=RS3("Apr4_P2")
va_apr5=RS3("Apr5_P2")
va_apr6=RS3("Apr6_P2")
va_apr7=RS3("Apr7_P2")
va_apr8=RS3("Apr8_P2")
va_v_apr1=RS3("V_Apr1_P2")
va_v_apr2=RS3("V_Apr2_P2")
va_v_apr3=RS3("V_Apr3_P2")
va_v_apr4=RS3("V_Apr4_P2")
va_v_apr5=RS3("V_Apr5_P2")
va_v_apr6=RS3("V_Apr6_P2")
va_v_apr7=RS3("V_Apr7_P2")
va_v_apr8=RS3("V_Apr8_P2")
va_sapr=RS3("VA_Sapr2")
va_pr=RS3("VA_Pr2")
va_te=RS3("VA_Te2")
va_bon=RS3("VA_Bon2")
va_me=RS3("VA_Me2")
va_mc=RS3("VA_Mc2")
va_faltas=RS3("NU_Faltas_P2")
usuario_grav= RS3("CO_Usuario")
data_grav=RS3("DA_Ult_Acesso")
hora_grav=RS3("HO_ult_Acesso")

va_pr=va_pr*1
va_te=va_te*1
pr=va_pr+va_te
elseif periodo_check=3 then
va_apr1=RS3("Apr1_P3")
va_apr2=RS3("Apr2_P3")
va_apr3=RS3("Apr3_P3")
va_apr4=RS3("Apr4_P3")
va_apr5=RS3("Apr5_P3")
va_apr6=RS3("Apr6_P3")
va_apr7=RS3("Apr7_P3")
va_apr8=RS3("Apr8_P3")
va_v_apr1=RS3("V_Apr1_P3")
va_v_apr2=RS3("V_Apr2_P3")
va_v_apr3=RS3("V_Apr3_P3")
va_v_apr4=RS3("V_Apr4_P3")
va_v_apr5=RS3("V_Apr5_P3")
va_v_apr6=RS3("V_Apr6_P3")
va_v_apr7=RS3("V_Apr7_P3")
va_v_apr8=RS3("V_Apr8_P3")
va_sapr=RS3("VA_Sapr3")
va_pr=RS3("VA_Pr3")
va_te=RS3("VA_Te3")
va_bon=RS3("VA_Bon3")
va_me=RS3("VA_Me3")
va_mc=RS3("VA_Mc3")
va_faltas=RS3("NU_Faltas_P3")
usuario_grav= RS3("CO_Usuario")
data_grav=RS3("DA_Ult_Acesso")
hora_grav=RS3("HO_ult_Acesso")

va_pr=va_pr*1
va_te=va_te*1
pr=va_pr+va_te
elseif periodo_check=4 then
va_apr1=RS3("Apr1_EC")
va_apr2=RS3("Apr2_EC")
va_apr3=RS3("Apr3_EC")
va_apr4=RS3("Apr4_EC")
va_apr5=RS3("Apr5_EC")
va_apr6=RS3("Apr6_EC")
va_apr7=RS3("Apr7_EC")
va_apr8=RS3("Apr8_EC")
va_v_apr1=RS3("V_Apr1_EC")
va_v_apr2=RS3("V_Apr2_EC")
va_v_apr3=RS3("V_Apr3_EC")
va_v_apr4=RS3("V_Apr4_EC")
va_v_apr5=RS3("V_Apr5_EC")
va_v_apr6=RS3("V_Apr6_EC")
va_v_apr7=RS3("V_Apr7_EC")
va_v_apr8=RS3("V_Apr8_EC")
va_sapr=RS3("VA_Sapr_EC")
va_pr=RS3("VA_Pr4")
va_me=RS3("VA_Me_EC")
va_mc=RS3("VA_Mfinal")
usuario_grav= RS3("CO_Usuario")
data_grav=RS3("DA_Ult_Acesso")
hora_grav=RS3("HO_ult_Acesso")


pr=va_pr
end if

if va_me="" or isnull(va_me)then
else
va_me=va_me/10
'	decimo = va_me - Int(va_me)
'		If decimo >= 0.5 Then
'			nota_arredondada = Int(va_me) + 1
'			va_me=nota_arredondada
'		Else
'			nota_arredondada = Int(va_me)
'			va_me=nota_arredondada					
'		End If
	va_me = formatNumber(va_me,1)
end if		
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
'		showapr="n"
'		else 
		showapr="s"
'		end if
'		if co_prova1="D"then
'		showprova="n"
'		else 
		showprova="s"
'		end if
'elseif periodo_check=2 then		
'		if co_apr2="D"then
'		showapr="n"
'		else 
		showapr="s"
'		end if
'		if co_prova2="D"then
'		showprova="n"
'		else 
		showprova="s"
'		end if					
'elseif periodo_check=3 then		
'		if co_apr3="D"then
'		showapr="n"
'		else 
		showapr="s"
'		end if
'		if co_prova3="D"then
'		showprova="n"
'		else 
		showprova="s"
'		end if
'elseif periodo_check=4 then		
'		if co_apr4="D"then
'		showapr="n"
'		else 
		showapr="s"
'		end if
'		if co_prova4="D"then
'		showprova="n"
'		else 
		showprova="s"
'		end if
'end if											
		
if hora_grav="nulo" then
hora_de=""
else
dados_hrd= split(hora_grav, ":" )
h_de= dados_hrd(0)
min_de= dados_hrd(1)
h_de=h_de*1
min_de=min_de*1


	if h_de<10 then
	h_de="0"&h_de
	end if
	if min_de<10 then
	min_de="0"&min_de
	end if	
	hora_de=h_de&":"&min_de
				
end if		
					
if data_grav="nulo"	then
data_inicio=""
else
		
dados_dtd= split(data_grav, "/" )
dia_de= dados_dtd(0)
mes_de= dados_dtd(1)
ano_de= dados_dtd(2)
dia_de=dia_de*1
mes_de=mes_de*1
if dia_de<10 then
dia_de="0"&dia_de
end if
if mes_de<10 then
mes_de="0"&mes_de
end if
data_inicio=dia_de&"/"&mes_de&"/"&ano_de&", "&hora_de
end if

				
end if
	if isnull(va_apr1) OR va_apr1="" OR isnull(va_apr2) OR va_apr2="" OR isnull(va_apr3) OR va_apr3="" OR isnull(va_apr4) OR va_apr4="" OR isnull(va_apr5) OR va_apr5="" OR isnull(va_apr6) OR va_apr6="" OR isnull(va_apr7) OR va_apr7="" OR isnull(va_apr8) OR va_apr8="" OR ISNULL(va_sapr) OR va_sapr="" then
	'data_inicio=""
	va_faltas=""
	elseif (va_apr1=0 OR va_apr1="0")and (va_apr2=0 OR va_apr2="0" )and (va_apr3=0 OR va_apr3="0" )and (va_apr4=0 OR va_apr4="0")and (va_apr5=0 OR va_apr5="0")and (va_apr6=0 OR va_apr6="0")and (va_apr7=0 OR va_apr7="0")and (va_apr8=0 OR va_apr8="0")and (va_sapr=0 OR va_sapr="0")  then
	'data_inicio=""
	va_faltas=""
	else
	data_inicio=data_inicio
	va_faltas=va_faltas
	end if

if usuario_grav="" or isnull(usuario_grav) then
no_usuario=""
else
		Set RS_pro = Server.CreateObject("ADODB.Recordset")
		SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
		RS_pro.Open SQL_pro, CON

no_usuario=RS_pro("NO_Usuario")
end if
%>
                        <tr class="<%response.Write(cor)%>"> 
                          <td width="<%response.Write(width)%>"
><font class="style1"> 
                            <%response.Write(no_materia)%>
                            </font></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showprova="n" AND showapr="n" then
							else							
							response.Write(va_faltas)
							End IF							
							%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showapr="s" then					
							response.Write(va_apr1)
							end if
							%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showapr="s" then					
							response.Write(va_v_apr1)
							end if
							%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showapr="s" then					
							response.Write(va_apr2)
							end if
							%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showapr="s" then					
							response.Write(va_v_apr2)
							end if
							%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showapr="s" then					
							response.Write(va_apr3)
							end if
							%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showapr="s" then					
							response.Write(va_v_apr3)
							end if
							%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showapr="s" then					
							response.Write(va_apr4)
							end if
							%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showapr="s" then					
							response.Write(va_v_apr4)
							end if
							%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showapr="s" then					
							response.Write(va_apr5)
							end if
							%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showapr="s" then					
							response.Write(va_v_apr5)
							end if
							%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showapr="s" then					
							response.Write(va_apr6)
							end if
							%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showapr="s" then					
							response.Write(va_v_apr6)
							end if
							%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showapr="s" then					
							response.Write(va_apr7)
							end if
							%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showapr="s" then					
							response.Write(va_v_apr7)
							end if
							%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showapr="s" then					
							response.Write(va_apr8)
							end if
							%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showapr="s" then					
							response.Write(va_v_apr8)
							end if
							%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showapr="s" then
	if va_sapr="" or isnull(va_sapr) then
	else						
	va_sapr = formatNumber(va_sapr,1)
	end if					
							response.Write(va_sapr)
							end if%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showprova="s" then
	if pr="" or isnull(pr) then
	else							
	pr = formatNumber(pr,1)												
							response.Write(pr)
	end if							
							end if%>
                              </font></div></td>
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showprova="s" then
							response.Write(va_me)							
							else					
							end if%>
                              </font></div></td>
<%if periodo_check=2 then%>							  
                          <td width="30"
> <div align="center"><font class="style1"> 
                              <%
							if showprova="s" then
							response.Write(va_bon)							
							else					
							end if%>
                              </font></div></td>
<%end if%>							  
                          <td width="170"
><div align="center"><font class="style1"> <%response.Write(no_usuario)%></font></div></td>
                          <td width="115"
> <div align="center"><font class="style1"> 
                              <%
'							if showprova="n" AND showapr="n" then
'							else							
							response.Write(data_inicio)
'							End if
							%>
                              </font></div></td>
                        </tr>
                        <%check=check+1
RSprog.MOVENEXT
wend%>
                        <tr valign="bottom"> 
<%if periodo_check=2 then%>							  
                          <td height="20" colspan="24" 
> <div align="right"><font class="style1"> F-Faltas , N-Nota Apr, P-Peso Apr, SAPR–Média 
                              das Aprs, PR-Prova, MP–Média Período e ECE1-1ª Etapa Complementar de Estudos</font></div></td>							  
<%else%>							
                          <td height="20" colspan="23" 
> <div align="right"><font class="style1"> F-Faltas , N-Nota Apr, P-Peso Apr, SAPR–Média 
                              das Aprs, PR-Prova e MP–Média Período</font></div></td>
<%end if%>						  
                        </tr>
                      </table>
                      <%end if%>
                    </td>
                  </tr>
                </table>
                <font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></div></td>
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
<script type="text/javascript">
<!--
  initInputHighlightScript();
//-->
</script>

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