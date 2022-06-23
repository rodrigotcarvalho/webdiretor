<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<% 
val_param = 70
opt= request.QueryString("opt")
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
ori = request.QueryString("ori")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
	

z = request.QueryString("z")
erro = request.QueryString("e")
vindo = request.QueryString("vd")
obr = request.QueryString("o")

if vindo="1" then
periodo_check=request.form("periodo")
cod= request.form("cod_cons")
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

if not RS.EOF then

ano_aluno = RS("NU_Ano")
rematricula = RS("DA_Rematricula")
situacao = RS("CO_Situacao")
encerramento= RS("DA_Encerramento")
unidade= RS("NU_Unidade")
curso= RS("CO_Curso")
etapa= RS("CO_Etapa")
turma= RS("CO_Turma")
cham= RS("NU_Chamada")
else
response.Write("Aluno n�o matriculado no ano letivo "&ano_letivo)
response.End()
end if

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
                    <input name="cod_cons" type="hidden" value="<%=codigo%>">
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
                      <select name="periodo" class="select_style" id="periodo" onChange="MM_callJS('submitfuncao()')">
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
                      <div align="center"><font class="form_dado_texto">
                        <%response.Write("<br><br><br><br><br>N�o existe Boletim de Avalia��es para este aluno!")%>
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
		
elseif notaFIL ="TB_NOTA_D" then
		CAMINHOn = CAMINHO_nd

elseif notaFIL ="TB_NOTA_E" then
		CAMINHOn = CAMINHO_ne	
				
elseif notaFIL ="TB_NOTA_F" then
		CAMINHOn = CAMINHO_nf	
				
elseif notaFIL ="TB_NOTA_K" then
		CAMINHOn = CAMINHO_nk	
						
elseif notaFIL ="TB_NOTA_L" then
		CAMINHOn = CAMINHO_nl	
		
elseif notaFIL ="TB_NOTA_M" then
		CAMINHOn = CAMINHO_nm	
										
elseif notaFIL ="TB_NOTA_V" then
		CAMINHOn = CAMINHO_nv			
				
else
		response.Write("ERRO")
end if	



		Set CON_N = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIRn

if notaFIL="TB_NOTA_A" then
minimo_recuperacao= 60
%>
                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            
                          <td width="125" class="tb_subtit"> 
                            <div align="left">Disciplina</div></td>						 
							
                           
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">T1</div></td>
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">T2</div></td>
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">T3</div></td>
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">T4</div></td>
							
                          <td width="31" class="tb_subtit"> 
                            <div align="center">MT</div></td>
<!--                          <td width="31" class="tb_subtit">&nbsp;</td>	-->						
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">PR1</div></td>
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">PR2</div></td>
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">PR3</div></td>
							
                          <td width="31" class="tb_subtit"> 
                            <div align="center"> MP</div></td>
<!--                          <td width="31" class="tb_subtit">&nbsp;</td>	-->					
							
                          <td width="31" class="tb_subtit"> 
                            <div align="center">M1</div></td>
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">Bon</div></td>
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">M2</div></td>
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">Rec</div></td>
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">M3</div></td>
                            
                          <td width="242" class="tb_subtit">
<div align="center">Alterado por</div></td>
                            <td width="115" class="tb_subtit"> <div align="center">Data/Hora</div></td>
                       </tr>
 <!--                       <tr>
                          <td width="125" class="tb_subtit">&nbsp;</td> 
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">N</div></td>
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">N</div></td>
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">N</div></td>
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">N</div></td>
							
                          <td width="31" class="tb_subtit"> 
                            <div align="center">M</div></td>
							
                          <td width="31" class="tb_subtit"> 
                            <div align="center">P</div></td>							
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">N</div></td>
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">N</div></td>
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">N</div></td>
							
                          <td width="31" class="tb_subtit"> 
                            <div align="center"> M</div></td>
                          <td width="31" class="tb_subtit"> 
                            <div align="center">P</div></td>							
							
                          <td width="31" class="tb_subtit"> 
                            <div align="center">M</div></td>
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">N</div></td>
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">M</div></td>
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">N</div></td>
							
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="176" class="tb_subtit">&nbsp;</td>
                          <td width="115" class="tb_subtit">&nbsp;</td>
                        </tr>-->
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
	
		Set RSnFIL = Server.CreateObject("ADODB.Recordset")
	  	Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' and NU_Periodo="&periodo_check
		Set RS3 = CON_N.Execute(SQL_N)
if RS3.EOF then
		va_pt="&nbsp;"
		va_pp="&nbsp;"
		va_t1="&nbsp;"
		va_t2="&nbsp;"
		va_t3="&nbsp;"
		va_t4="&nbsp;"
		va_mt="&nbsp;"
		va_p1="&nbsp;"
		va_p2="&nbsp;"
		va_p3="&nbsp;"
		va_mp="&nbsp;"
		va_m1="&nbsp;"
		va_bon="&nbsp;"
		va_m2="&nbsp;"
		va_rec="&nbsp;"
		va_m3="&nbsp;"
		data_grav="&nbsp;"
		hora_grav="&nbsp;"
		usuario_grav="&nbsp;"			
else
		va_pt=RS3("PE_Teste")
		va_pp=RS3("PE_Prova")
		va_t1=RS3("VA_Teste1")
		va_t2=RS3("VA_Teste2")
		if notaFIL<>"TB_NOTA_E" then
			va_t3=RS3("VA_Teste3")
			va_t4=RS3("VA_Teste4")	
		end if
		va_mt=RS3("MD_Teste")
		va_p1=RS3("VA_Prova1")
		va_p2=RS3("VA_Prova2")
		va_p3=RS3("VA_Prova3")
		va_mp=RS3("MD_Prova")
		va_m1=RS3("VA_Media1")
		va_bon=RS3("VA_Bonus")
		va_m2=RS3("VA_Media2")
		va_rec=RS3("VA_Rec")
		va_m3=RS3("VA_Media3")
		data_grav=RS3("DA_Ult_Acesso")
		hora_grav=RS3("HO_ult_Acesso")
		usuario_grav=RS3("CO_Usuario")
end if

									
		
if hora_grav="&nbsp;" then
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
					
if data_grav="&nbsp;" then
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

showapr="s"
showprova="s"
'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
'data_inicio=""
'va_faltas=""
'		end if

if usuario_grav="&nbsp;" then
no_usuario=""
else
		Set RS_pro = Server.CreateObject("ADODB.Recordset")
		SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
		RS_pro.Open SQL_pro, CON

		if RS_pro.eof then
			no_usuario= usuario_grav	
		else
			no_usuario=RS_pro("NO_Usuario")
		end if	
end if
%>
                        <tr class="<%response.Write(cor)%>"> 
                          <td width="125">
                            <%response.Write(no_materia)%>
                          </td>
                          <td width="31"> 
                            <div align="center">
                              <%
							if showapr="s" then							
							response.Write(va_t1)
							End IF							
							%>
                            </div></td>
                          <td width="31"> 
                            <div align="center">
                              <%
							if showapr="s" then					
							response.Write(va_t2)
							end if
							%>
                            </div></td>
                          <td width="31"
> 
                            <div align="center">
                              <%
							if showapr="s" then					
							response.Write(va_t3)
							end if
							%>
                            </div></td>
                          <td width="31"
> 
                            <div align="center">
                              <%
							if showapr="s" then					
							response.Write(va_t4)
							end if
							%>
                            </div></td>
                          <td width="31"
> 
                            <div align="center">
                              <%
							if showapr="s" then					
							response.Write(va_mt)
							end if
							%>
                            </div></td>
<!--                          <td width="31"
> 
                            <div align="center">
                              <%
							if showapr="s" then					
							'response.Write(va_pt)
							end if
							%>
                            </div></td>	-->						  
                          <td width="31"
> 
                            <div align="center"> 
                              <%
							if showprova="s" then					
							response.Write(va_p1)
							end if
							%>
                            </div></td>
                          <td width="31"
> 
                            <div align="center"> 
                              <%
							if showprova="s" then					
							response.Write(va_p2)
							end if
							%>
                            </div></td>
                          <td width="31"
> 
                            <div align="center"> 
                              <%
							if showprova="s" then					
							response.Write(va_p3)
							end if
							%>
                            </div></td>
                          <td width="31"
> 
                            <div align="center"> 
                              <%
							if showprova="s" then					
							response.Write(va_mp)
							end if
							%>
                            </div></td>
 <!--                         <td width="31"
> 
                            <div align="center"> 
                              <%
							if showprova="s" then					
							'response.Write(va_pp)
							end if
							%>
                            </div></td>-->							  
                          <td width="31"
> 
                            <div align="center"> 
                              <%
							if showprova="s" AND showapr="s" then	
								if isnumeric(va_m1) then
								    va_m1 = va_m1*1
									val_param = val_param*1
									if va_m1<val_param then
										response.Write("<font color=#F00>"& va_m1&"</font>")	
									else
										response.Write(va_m1)									
									end if									
								else	
									response.Write(va_m1)
								end if	
							end if
							%>
                            </div></td>
                          <td width="31"
> 
                            <div align="center"> 
                              <%
							if showprova="s" AND showapr="s" then					
							response.Write(va_bon)
							end if
							%>
                            </div></td>
                          <td width="31"
> 
                            <div align="center"> 
                              <%
							if showprova="s" AND showapr="s" then
								if isnumeric(va_m2) then
								    va_m2 = va_m2*1
									val_param = val_param*1
									if va_m2<val_param then
										response.Write("<font color=#F00>"& va_m2&"</font>")	
									else
										response.Write(va_m2)									
									end if									
								else	
									response.Write(va_m2)
								end if													

							end if
							%>
                            </div></td>
                          <td width="31"
> 
                            <div align="center"> 
                              <%
							if showprova="s" AND showapr="s" then					
							response.Write(va_rec)
							end if
							%>
                            </div></td>
                          <td width="31"
> 
                            <div align="center"> 
                              <%
							if showprova="s" AND showapr="s" then					
								if isnumeric(va_m3) then
								    va_m3 = va_m3*1
									val_param = val_param*1
									if va_m3<val_param then
										response.Write("<font color=#F00>"& va_m3&"</font>")	
									else
										response.Write(va_m3)									
									end if									
								else	
									response.Write(va_m3)
								end if	
							end if
							%>
                            </div></td>
                          <td width="242"
>
<div align="center">							
							<%
							if showprova="s" AND showapr="s" then
							response.Write(no_usuario)
  							end if
 							%>
</div></td>
                          <td width="115"
> <div align="center"> 
                              <%
							if showprova="s" AND showapr="s" then							
							response.Write(data_inicio)
							End if
							%>
                              </div></td>
                        </tr>
                        <%check=check+1
RSprog.MOVENEXT
wend
%>
						                        <tr valign="bottom"> 
                          <td height="20" colspan="23" 
> <div align="right"><font class="form_corpo"> T-Teste, MT�M�dia dos Testes, PR-Prova, 
                              MP�M�dia das Provas, N-Nota, M-M&eacute;dia </font></div></td>
                        </tr>
                      </table>

<%
elseif notaFIL="TB_NOTA_B" or notaFIL="TB_NOTA_E" then
%>
                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td width="125" class="tb_subtit"> <div align="left">Disciplina</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">T1</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">T2</div></td>
                          <td width="31"  class="tb_subtit"><div align="center">T3</div></td>
                          <td width="31"  class="tb_subtit"><div align="center">T4</div></td>
                          <td width="31" class="tb_subtit"> <div align="center">MT</div></td>
<!--                          <td width="37" class="tb_subtit">&nbsp;</td>-->
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">PR1</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">S</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">PR2</div></td>
                          <td width="31" class="tb_subtit"> <div align="center"> 
                              MP</div></td>
<!--                          <td width="37" class="tb_subtit">&nbsp;</td>-->
                          <td width="31" class="tb_subtit"> 
                            <div align="center">M1</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">Bon</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">M2</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">Rec</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">M3</div></td>
                          <td width="242" class="tb_subtit"> <div align="center">Alterado 
                              por</div></td>
                          <td width="115" class="tb_subtit"> <div align="center">Data/Hora</div></td>
                        </tr>
<!--                        <tr>
                          <td width="125" class="tb_subtit">&nbsp;</td> 
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">P</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center"> 
                              M</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">P</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37"  class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37"  class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="176" class="tb_subtit">&nbsp;</td>
                          <td width="115" class="tb_subtit">&nbsp;</td>
                        </tr>
-->                        <%
		rec_lancado="sim"
		
				Set RSprog = Server.CreateObject("ADODB.Recordset")
				SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
				RSprog.Open SQLprog, CON0
		
		check=1
			
		while not RSprog.EOF
		
				materia=RSprog("CO_Materia")
				mae=RSprog("IN_MAE")
				fil=RSprog("IN_FIL")
				in_co=RSprog("IN_CO")
				nu_peso=RSprog("NU_Peso")
				ordem=RSprog("NU_Ordem_Boletim")
				
			peso_acumula=0
			m1_ac=0
			m2_ac=0			
			m3_ac=0
			m1_exibe=0
			m2_exibe=0
			m3_exibe=0
					
			if mae=TRUE THEN
			
			check=check+1
			
					Set RS1a = Server.CreateObject("ADODB.Recordset")
					SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia_Principal='"&materia&"'"
					RS1a.Open SQL1a, CON0
					
					if RS1a.EOF then
			
						Set RS1b = Server.CreateObject("ADODB.Recordset")
						SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
						RS1b.Open SQL1b, CON0
						
						no_materia=RS1b("NO_Materia")
						
						 if check mod 2 =0 then
						  cor = "tb_fundo_linha_par" 
						 else cor ="tb_fundo_linha_impar"
						 end if
							
						Set RSnFIL = Server.CreateObject("ADODB.Recordset")
						Set RS3 = Server.CreateObject("ADODB.Recordset")
						SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' and NU_Periodo="&periodo_check
						Set RS3 = CON_N.Execute(SQL_N)
						
						if RS3.EOF then
								va_pt="&nbsp;"
								va_pp="&nbsp;"
								va_t1="&nbsp;"
								va_t2="&nbsp;"
								va_t3="&nbsp;"
								va_t4="&nbsp;"								
								va_mt="&nbsp;"
								va_p1="&nbsp;"
								va_p2="&nbsp;"
								va_p3="&nbsp;"
								va_mp="&nbsp;"
								va_m1="&nbsp;"
								va_bon="&nbsp;"
								va_m2="&nbsp;"
								va_rec="&nbsp;"
								va_m3="&nbsp;"
								data_grav="&nbsp;"
								hora_grav="&nbsp;"
								usuario_grav="&nbsp;"			
						else
								va_pt=RS3("PE_Teste")
								va_pp=RS3("PE_Prova")
								va_t1=RS3("VA_Teste1")
								va_t2=RS3("VA_Teste2")

								if notaFIL<>"TB_NOTA_E" then
									va_t3=RS3("VA_Teste3")
									va_t4=RS3("VA_Teste4")	
								end if								
								va_mt=RS3("MD_Teste")
								va_p1=RS3("VA_Prova1")
								va_p2=RS3("VA_Simul")
								va_p3=RS3("VA_Prova2")
								va_mp=RS3("MD_Prova")
								va_m1=RS3("VA_Media1")
								va_bon=RS3("VA_Bonus")
								va_m2=RS3("VA_Media2")
								va_rec=RS3("VA_Rec")
								va_m3=RS3("VA_Media3")
								data_grav=RS3("DA_Ult_Acesso")
								hora_grav=RS3("HO_ult_Acesso")
								usuario_grav=RS3("CO_Usuario")
						end if
						
															
								
						if hora_grav="&nbsp;" then
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
											
						if data_grav="&nbsp;" then
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
							
							showapr="s"
							showprova="s"
						'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
						'data_inicio=""
						'va_faltas=""
						'		end if
						
						if usuario_grav="&nbsp;" then
							no_usuario=""
						else
								Set RS_pro = Server.CreateObject("ADODB.Recordset")
								SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
								RS_pro.Open SQL_pro, CON
							
							if RS_pro.eof then
								no_usuario= usuario_grav	
							else
								no_usuario=RS_pro("NO_Usuario")
							end if	
						end if
						%>
												<tr class="<%response.Write(cor)%>"> 
												  <td width="125"> 
													<%response.Write(no_materia)%></td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then							
													response.Write(va_t1)
													End IF							
													%>
													</div></td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_t2)
													end if
													%>
													</div></td>
												  <td width="31">
                                                  <div align="center">
												  <%
													if showapr="s" then					
													response.Write(va_t3)
													end if
													%></div></td>
												  <td width="31">
                                                  <div align="center">
                                                 <%
													if showapr="s" then					
													response.Write(va_t4)
													end if
													%></div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_mt)
													end if
													%>
													</div></td>
<!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													'response.Write(va_pt)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													response.Write(va_p1)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													response.Write(va_p2)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													response.Write(va_p3)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													response.Write(va_mp)
													end if
													%>
													</div></td>
<!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													'response.Write(va_pp)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
																					if isnumeric(va_m1) then
									if va_m1<val_param then
										response.Write("<font color=#F00>"& va_m1&"</font>")	
									else
										response.Write(va_m1)									
									end if									
								else	
									response.Write(va_m1)
								end if	

													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_bon)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
								if isnumeric(va_m2) then
								    va_m2 = va_m2*1
									val_param = val_param*1
									if va_m2<val_param then
										response.Write("<font color=#F00>"& va_m2&"</font>")	
									else
										response.Write(va_m2)									
									end if									
								else	
									response.Write(va_m2)
								end if	
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_rec)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
								if isnumeric(va_m3) then
								    va_m3 = va_m3*1
									val_param = val_param*1
									if va_m3<val_param then
										response.Write("<font color=#F00>"& va_m3&"</font>")	
									else
										response.Write(va_m3)									
									end if									
								else	
									response.Write(va_m3)
								end if	
													end if
													%>
													</div></td>
												  <td width="242"
						> <div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then
													response.Write(no_usuario)
													end if
													%>
													</div></td>
												  <td width="115"
						> <div align="center"> 
													<%
													if showprova="s" AND showapr="s" then							
													response.Write(data_inicio)
													End if
													%></div>
												  </td>
												</tr>
					<%
			else

			
			
				 if check mod 2 =0 then
				  cor = "tb_fundo_linha_par" 
				 else cor ="tb_fundo_linha_impar"
				  end if
			
						Set RS1b = Server.CreateObject("ADODB.Recordset")
						SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
						RS1b.Open SQL1b, CON0
						
				no_materia=RS1b("NO_Materia")
					
						Set RSnFIL = Server.CreateObject("ADODB.Recordset")
						Set RS3 = Server.CreateObject("ADODB.Recordset")
						SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' and NU_Periodo="&periodo_check
						Set RS3 = CON_N.Execute(SQL_N)
				if RS3.EOF then
						va_pt="&nbsp;"
						va_pp="&nbsp;"
						va_t1="&nbsp;"
						va_t2="&nbsp;"
						va_t3="&nbsp;"
						va_t4="&nbsp;"						
						va_mt="&nbsp;"
						va_p1="&nbsp;"
						va_p2="&nbsp;"
						va_p3="&nbsp;"
						va_mp="&nbsp;"
						va_m1="&nbsp;"
						va_bon="&nbsp;"
						va_m2="&nbsp;"
						va_rec="&nbsp;"
						va_m3="&nbsp;"
						data_grav="&nbsp;"
						hora_grav="&nbsp;"
						usuario_grav="&nbsp;"			
				else
						va_pt=RS3("PE_Teste")
						va_pp=RS3("PE_Prova")
						va_t1=RS3("VA_Teste1")
						va_t2=RS3("VA_Teste2")
						if notaFIL<>"TB_NOTA_E" then
							va_t3=RS3("VA_Teste3")
							va_t4=RS3("VA_Teste4")	
						end if
						
						va_mt=RS3("MD_Teste")
						va_p1=RS3("VA_Prova1")
						va_p2=RS3("VA_Simul")
						va_p3=RS3("VA_Prova2")
						va_mp=RS3("MD_Prova")
						va_m1=RS3("VA_Media1")
						va_bon=RS3("VA_Bonus")
						va_m2=RS3("VA_Media2")
						va_rec=RS3("VA_Rec")
						va_m3=RS3("VA_Media3")
						data_grav=RS3("DA_Ult_Acesso")
						hora_grav=RS3("HO_ult_Acesso")
						usuario_grav=RS3("CO_Usuario")
				end if

						
				if hora_grav="&nbsp;" then
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
									
				if data_grav="&nbsp;" then
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
				
				showapr="s"
				showprova="s"
				'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
				'data_inicio=""
				'va_faltas=""
				'		end if
				
				if usuario_grav="&nbsp;" then
				no_usuario=""
				else
						Set RS_pro = Server.CreateObject("ADODB.Recordset")
						SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
						RS_pro.Open SQL_pro, CON
				
							if RS_pro.eof then
								no_usuario= usuario_grav	
							else
								no_usuario=RS_pro("NO_Usuario")
							end if	
				end if
				%>
										<tr class="<%response.Write(cor)%>"> 
										  <td width="125"> 
											<%response.Write(no_materia)%></td>
										  <td width="31"> 
											<div align="center"> 
											  <%
											if showapr="s" then							
											response.Write(va_t1)
											End IF							
											%>
											</div></td>
										  <td width="31"> 
											<div align="center"> 
											  <%
											if showapr="s" then					
											response.Write(va_t2)
											end if
											%>
											</div></td>
										  <td width="31"><div align="center">
										    <%
													if showapr="s" then					
													response.Write(va_t3)
													end if
													%>
										    </div></td>
										  <td width="31"><div align="center">
										    <%
													if showapr="s" then					
													response.Write(va_t4)
													end if
													%>
										    </div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showapr="s" then					
											response.Write(va_mt)
											end if
											%>
											</div></td>
<!--										  <td width="37"
				> 
											<div align="center"> 
											  <%
											if showapr="s" then					
											'response.Write(va_pt)
											end if
											%>
											</div></td>-->
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" then					
											response.Write(va_p1)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" then					
											response.Write(va_p2)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" then					
											response.Write(va_p3)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" and materia<>"LP" then					
											response.Write(va_mp)
											end if
											%>
											</div></td>
<!--										  <td width="37"
				> 
											<div align="center"> 
											  <%
											if showprova="s" then					
											'response.Write(va_pp)
											end if
											%>
											</div></td>-->
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_m1)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_bon)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_m2)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_rec)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_m3)
											end if
											%>
											</div></td>
										  <td width="242"
				> <div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then
											response.Write(no_usuario)
											end if
											%>
											</div></td>
										  <td width="115"
				> <div align="center"> 
											<%
											if showprova="s" AND showapr="s" then							
											response.Write(data_inicio)
											End if
											%></div>
										  </td>
										</tr>
			<%
			peso_acumula=0
			acumula_m1=0
			m2_ac=0			
			m3_ac=0
			m1_exibe=0
			m2_exibe=0
			m3_exibe=0
			
			
				while not RS1a.EOF
				
						materia_fil=RS1a("CO_Materia")
					
								Set RS1b = Server.CreateObject("ADODB.Recordset")
								SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia_fil&"'"
								RS1b.Open SQL1b, CON0
								
						no_materia_fil=RS1b("NO_Materia")
						
						Set RSpa = Server.CreateObject("ADODB.Recordset")
						SQLpa= "SELECT * FROM TB_Programa_Aula where CO_Materia='"&materia_fil&"' AND CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"'"
						RSpa.Open SQLpa, CON0
												
						nu_peso_fil=RSpa("NU_Peso")			
						
						if isnull(nu_peso_fil) or nu_peso_fil="" then
							nu_peso_fil=1
						end if			
							
								Set RSnFIL = Server.CreateObject("ADODB.Recordset")
								Set RS3 = Server.CreateObject("ADODB.Recordset")
								SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia_Principal = '"& materia &"' AND CO_Materia = '"& materia_fil &"' and NU_Periodo="&periodo_check
								Set RS3 = CON_N.Execute(SQL_N)
						if RS3.EOF then
								va_pt="&nbsp;"
								va_pp="&nbsp;"
								va_t1="&nbsp;"
								va_t2="&nbsp;"
								va_t3="&nbsp;"
								va_t4="&nbsp;"									
								va_mt="&nbsp;"
								va_p1="&nbsp;"
								va_p2="&nbsp;"
								va_p3="&nbsp;"
								va_mp="&nbsp;"
								va_m1=0
								va_bon="&nbsp;"
								va_m2=0
								va_rec="&nbsp;"
								va_m3=0
								data_grav="&nbsp;"
								hora_grav="&nbsp;"
								usuario_grav="&nbsp;"			
						else
								va_pt=RS3("PE_Teste")
								va_pp=RS3("PE_Prova")
								va_t1=RS3("VA_Teste1")
								va_t2=RS3("VA_Teste2")
								if notaFIL<>"TB_NOTA_E" then
									va_t3=RS3("VA_Teste3")
									va_t4=RS3("VA_Teste4")	
								end if								
								va_mt=RS3("MD_Teste")
								va_p1=RS3("VA_Prova1")
								va_p2=RS3("VA_Simul")
								va_p3=RS3("VA_Prova2")
								va_mp=RS3("MD_Prova")
								va_m1=RS3("VA_Media1")
								va_bon=RS3("VA_Bonus")
								va_m2=RS3("VA_Media2")
								va_rec=RS3("VA_Rec")
								va_m3=RS3("VA_Media3")
								data_grav=RS3("DA_Ult_Acesso")
								hora_grav=RS3("HO_ult_Acesso")
								usuario_grav=RS3("CO_Usuario")
						end if
						if isnull(va_m1) or va_m1="" then
							'va_m1=0
							sem_media1="s"
						else
							sem_media1="n"
							m1_ac=m1_ac+(va_m1*nu_peso_fil)															
						end if

						if isnull(va_m2) or va_m2="" then
							'va_m2=0
							sem_media2="s"
						else
							sem_media2="n"	
							m2_ac=m2_ac+(va_m2*nu_peso_fil)							
						end if
						
						if isnull(va_m3) or va_m3="" then
							'va_m3=0
							sem_media3="s"
						else
							sem_media3="n"		
							m3_ac=m3_ac+(va_m3*nu_peso_fil)					
						end if												
						
							peso_acumula=peso_acumula+nu_peso_fil
													
								
						if hora_grav="&nbsp;" then
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
											
						if data_grav="&nbsp;" then
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
						
						showapr="s"
						showprova="s"
						'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
						'data_inicio=""
						'va_faltas=""
						'		end if
						
						if usuario_grav="&nbsp;" then
						no_usuario=""
						else
								Set RS_pro = Server.CreateObject("ADODB.Recordset")
								SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
								RS_pro.Open SQL_pro, CON
						
							if RS_pro.eof then
								no_usuario= usuario_grav	
							else
								no_usuario=RS_pro("NO_Usuario")
							end if	
						end if
						%>
												
											<tr class="<%response.Write(cor)%>"> 
											  <td width="125">&nbsp;&nbsp;&nbsp;
												  <%response.Write(no_materia_fil)%>
											  </td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then							
													response.Write(va_t1)
													End IF							
													%>
													</div></td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_t2)
													end if
													%>
													</div></td>
												  <td width="31"><div align="center">
												    <%
													if showapr="s" then					
													response.Write(va_t3)
													end if
													%>
										      </div></td>
												  <td width="31"><div align="center">
												    <%
													if showapr="s" then					
													response.Write(va_t4)
													end if
													%>
										      </div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_mt)
													end if
													%>
													</div></td>
<!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													'response.Write(va_pt)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													response.Write(va_p1)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													response.Write(va_p2)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													response.Write(va_p3)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" and materia<>"LP" then					
													response.Write(va_mp)
													end if
													%>
													</div></td>
<!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													'response.Write(va_pp)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m1)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_bon)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m2)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_rec)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m3)
													end if
													%>
													</div></td>
												  <td width="242"
						> <div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then
													response.Write(no_usuario)
													end if
													%>
													</div></td>
												  <td width="115"
						> <div align="center"> 
													<%
													if showprova="s" AND showapr="s" then							
													response.Write(data_inicio)
													End if
													%></div>
											  </td>
						</tr>
				<%
				RS1a.movenext
				wend
						if	sem_media1="s" then
							m1_exibe="&nbsp;"
						else
							m1_exibe=m1_ac/peso_acumula
							decimo = m1_exibe - Int(m1_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m1_exibe) + 1
									m1_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m1_exibe)
									m1_exibe=nota_arredondada					
								End If
							m1_exibe= formatNumber(m1_exibe,0)							
						end if	
							
						if	sem_media2="s" then	
							m2_exibe="&nbsp;"
						else												
							m2_exibe=m2_ac/peso_acumula
							decimo = m2_exibe - Int(m2_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m2_exibe) + 1
									m2_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m2_exibe)
									m2_exibe=nota_arredondada					
								End If
							m2_exibe= formatNumber(m2_exibe,0)							
						end if	
						
						if	sem_media3="s" then
							m3_exibe="&nbsp;"
						else							
							m3_exibe=m3_ac/peso_acumula
								decimo = m3_exibe - Int(m3_exibe)
									If decimo >= 0.5 Then
										nota_arredondada = Int(m3_exibe) + 1
										m3_exibe=nota_arredondada
									Else
										nota_arredondada = Int(m3_exibe)
										m3_exibe=nota_arredondada					
									End If
								m3_exibe= formatNumber(m3_exibe,0)									
						end if														
				
				%>
									<tr class="tb_fundo_linha_media"> 
									  <td width="125">&nbsp;&nbsp;&nbsp; M&eacute;dia</td>
										  <td width="31"> 
											<div align="center"></div></td>
										  <td width="31"> 
											<div align="center"> </div></td>
										  <td width="31"
				>&nbsp;</td>
										  <td width="31"
				>&nbsp;</td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
<!--										  <td width="37"
				> 
											<div align="center"> </div></td>-->
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
<!--										  <td width="37"
				> 
											<div align="center"> </div></td>-->
										  <td width="31"
				> 
											<div align="center">
											
											
											<%								if isnumeric(m1_exibe) then
									m1_exibe = m1_exibe*1	
									val_param = val_param*1												
									if m1_exibe<val_param then
										response.Write("<font color=#F00>"& m1_exibe&"</font>")	
									else
										response.Write(m1_exibe)									
									end if									
								else	
									response.Write(m1_exibe)
								end if	%> </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center">
                              <%if isnumeric(m2_exibe) then 
									m2_exibe = m2_exibe*1	
									val_param = val_param*1	
									if m2_exibe<val_param then
										response.Write("<font color=#F00>"& m2_exibe&"</font>")	
									else
										response.Write(m2_exibe)									
									end if									
								else	
									response.Write(m2_exibe)
								end if	%>
                            </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center">
                              <%if isnumeric(m3_exibe) then 
									m3_exibe = m3_exibe*1	
									val_param = val_param*1	
									if m3_exibe<val_param then
										response.Write("<font color=#F00>"& m3_exibe&"</font>")	
									else
										response.Write(m3_exibe)									
									end if									
								else	
									response.Write(m3_exibe)
								end if	%>
                            </div></td>
										  <td width="242"
				> <div align="center"> </div></td>
										  <td width="115"
				> <div align="center"> </div>
									  </td>
						</tr>
			<%
			end if
			end if

		RSprog.MOVENEXT
		wend
		%>
								<tr valign="bottom"> 
								  <td height="20" colspan="23" 
		> <div align="right"><font class="form_corpo"> 
        
        <% if etapa=6 or etapa=7 or etapa=8 or etapa=9 then
				Response.Write("T-Teste, PR-Prova, S - Simulado, MT=(T1+T2+T3+T4)/4, MP=(PR1+S), M3=((MTx1)+(MPx2))/3. <br>Para a Disciplina Portugu�s PR2 = Reda��o e M3=((MTx1)+(MPx2)+(PR2x2))/5.")			
		else
			Response.Write("T-Teste, MT�M�dia dos Testes, PR-Prova, MP�M�dia das Provas e M-M&eacute;dia")
		End if%>        
        
        </font></div></td>
								</tr>
					  </table>
<%
elseif notaFIL="TB_NOTA_C" then
%>
                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td width="125" class="tb_subtit"> <div align="left">Disciplina</div></td>
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">T1</div></td>
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">T2</div></td>
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">T3</div></td>
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">T4</div></td>
                          <td width="33" class="tb_subtit"> <div align="center">MT</div></td>
<!--                          <td width="33" class="tb_subtit">&nbsp;</td>-->
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">PR1</div></td>
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">PR2</div></td>
                          <td width="33" class="tb_subtit"> <div align="center"> 
                              MP</div></td>
<!--                          <td width="33" class="tb_subtit">&nbsp;</td>-->
                          <td width="33" class="tb_subtit"> 
                            <div align="center">M1</div></td>
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">Bon</div></td>
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">M2</div></td>
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">Rec</div></td>
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">M3</div></td>
                          <td width="243" class="tb_subtit"> 
                            <div align="center">Alterado 
                              por</div></td>
                          <td width="115" class="tb_subtit"> <div align="center">Data/Hora</div></td>
                        </tr>
<!--                        <tr>
                          <td width="125" class="tb_subtit">&nbsp;</td> 
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="33" class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="33" class="tb_subtit"> 
                            <div align="center">P</div></td>
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="33" class="tb_subtit"> 
                            <div align="center"> 
                              M</div></td>
                          <td width="33" class="tb_subtit"> 
                            <div align="center">P</div></td>
                          <td width="33" class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="33"  class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="177" class="tb_subtit">&nbsp;</td>
                          <td width="115" class="tb_subtit">&nbsp;</td>
                        </tr>
-->                        <%
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
	
		Set RSnFIL = Server.CreateObject("ADODB.Recordset")
	  	Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' and NU_Periodo="&periodo_check
		Set RS3 = CON_N.Execute(SQL_N)
if RS3.EOF then
		va_pt="&nbsp;"
		va_pp="&nbsp;"
		va_t1="&nbsp;"
		va_t2="&nbsp;"
		va_t3="&nbsp;"
		va_t4="&nbsp;"
		va_mt="&nbsp;"
		va_p1="&nbsp;"
		va_p2="&nbsp;"
		va_mp="&nbsp;"
		va_m1="&nbsp;"
		va_bon="&nbsp;"
		va_m2="&nbsp;"
		va_rec="&nbsp;"
		va_m3="&nbsp;"
		data_grav="&nbsp;"
		hora_grav="&nbsp;"
		usuario_grav="&nbsp;"			
else
		va_pt=RS3("PE_Teste")
		va_pp=RS3("PE_Prova")
		va_t1=RS3("VA_Teste1")
		va_t2=RS3("VA_Teste2")
		if notaFIL<>"TB_NOTA_E" then
			va_t3=RS3("VA_Teste3")
			va_t4=RS3("VA_Teste4")	
		end if
		va_mt=RS3("MD_Teste")
		va_p1=RS3("VA_Prova1")
		va_p2=RS3("VA_Prova2")
		va_mp=RS3("MD_Prova")
		va_m1=RS3("VA_Media1")
		va_bon=RS3("VA_Bonus")
		va_m2=RS3("VA_Media2")
		va_rec=RS3("VA_Rec")
		va_m3=RS3("VA_Media3")
		data_grav=RS3("DA_Ult_Acesso")
		hora_grav=RS3("HO_ult_Acesso")
		usuario_grav=RS3("CO_Usuario")
end if

									
		
if hora_grav="&nbsp;" then
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
					
if data_grav="&nbsp;" then
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

showapr="s"
showprova="s"
'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
'data_inicio=""
'va_faltas=""
'		end if

if usuario_grav="&nbsp;" then
no_usuario=""
else
		Set RS_pro = Server.CreateObject("ADODB.Recordset")
		SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
		RS_pro.Open SQL_pro, CON

		if RS_pro.eof then
			no_usuario= usuario_grav	
		else
			no_usuario=RS_pro("NO_Usuario")
		end if	
end if
%>
                        <tr class="<%response.Write(cor)%>"> 
                          <td width="125"> 
                            <%response.Write(no_materia)%>
                          </td>
                          <td width="33"> 
                            <div align="center"> 
                              <%
							if showapr="s" then							
							response.Write(va_t1)
							End IF							
							%>
                            </div></td>
                          <td width="33"> 
                            <div align="center"> 
                              <%
							if showapr="s" then					
							response.Write(va_t2)
							end if
							%>
                            </div></td>
                          <td width="33"
> 
                            <div align="center"> 
                              <%
							if showapr="s" then					
							response.Write(va_t3)
							end if
							%>
                            </div></td>
                          <td width="33"
> 
                            <div align="center"> 
                              <%
							if showapr="s" then					
							response.Write(va_t4)
							end if
							%>
                            </div></td>
                          <td width="33"
> 
                            <div align="center"> 
                              <%
							if showapr="s" then					
							response.Write(va_mt)
							end if
							%>
                            </div></td>
<!--                          <td width="33"
> 
                            <div align="center"> 
                              <%
							if showapr="s" then					
							'response.Write(va_pt)
							end if
							%>
                            </div></td>-->
                          <td width="33"
> 
                            <div align="center"> 
                              <%
							if showprova="s" then					
							response.Write(va_p1)
							end if
							%>
                            </div></td>
                          <td width="33"
> 
                            <div align="center"> 
                              <%
							if showprova="s" then					
							response.Write(va_p2)
							end if
							%>
                            </div></td>
                          <td width="33"
> 
                            <div align="center"> 
                              <%if showprova="s" then					
													response.Write(va_bat)
													end if
													%>
                              <%
							if showprova="s" then					
							response.Write(va_mp)
							end if
							%>
                            </div></td>
<!--                          <td width="33"
> 
                            <div align="center"> 
                              <%
							if showprova="s" then					
						'	response.Write(va_pp)
							end if
							%>
                            </div></td>-->
                          <td width="33"
> 
                            <div align="center"> 
                              <%
							if showprova="s" AND showapr="s" then					
								if isnumeric(va_m1) then
									if va_m1<val_param then
										response.Write("<font color=#F00>"& va_m1&"</font>")	
									else
										response.Write(va_m1)									
									end if									
								else	
									response.Write(va_m1)
								end if	

							end if
							%>
                            </div></td>
                          <td width="33"
> 
                            <div align="center"> 
                              <%
							if showprova="s" AND showapr="s" then					
							response.Write(va_bon)
							end if
							%>
                            </div></td>
                          <td width="33"
> 
                            <div align="center"> 
                              <%
							if showprova="s" AND showapr="s" then					
								if isnumeric(va_m2) then
								    va_m2 = va_m2*1
									val_param = val_param*1
									if va_m2<val_param then
										response.Write("<font color=#F00>"& va_m2&"</font>")	
									else
										response.Write(va_m2)									
									end if									
								else	
									response.Write(va_m2)
								end if	
							end if
							%>
                            </div></td>
                          <td width="33"
> 
                            <div align="center"> 
                              <%
							if showprova="s" AND showapr="s" then					
							response.Write(va_rec)
							end if
							%>
                            </div></td>
                          <td width="33"
> 
                            <div align="center"> 
                              <%
							if showprova="s" AND showapr="s" then					
								if isnumeric(va_m3) then
								    va_m3 = va_m3*1
									val_param = val_param*1
									if va_m3<val_param then
										response.Write("<font color=#F00>"& va_m3&"</font>")	
									else
										response.Write(va_m3)									
									end if									
								else	
									response.Write(va_m3)
								end if	
							end if
							%>
                            </div></td>
                          <td width="243"
> 
                            <div align="center"> 
                              <%
							if showprova="s" AND showapr="s" then
							response.Write(no_usuario)
  							end if
 							%>
                            </div></td>
                          <td width="115"
> <div align="center"> 
                            <%
							if showprova="s" AND showapr="s" then							
							response.Write(data_inicio)
							End if
							%></div>
                          </td>
                        </tr>
                        <%check=check+1
RSprog.MOVENEXT
wend
%>
                        <tr valign="bottom"> 
                          <td height="20" colspan="22" 
> <div align="right"><font class="form_corpo"> T-Teste, MT�Soma dos Testes, PR-Prova, 
                                MP = (P1 + P2) / 2  , N-Nota, M-M&eacute;dia e   M3 = ((MTx1)+(MPx1)) / 2  </font></div></td>
                        </tr>
                      </table>
<%
elseif notaFIL="TB_NOTA_F" then
%>
                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td width="125" class="tb_subtit"> <div align="left">Disciplina</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">TD1</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">TD2</div></td>
                          <td width="31" class="tb_subtit"> <div align="center">MT</div></td>
<!--                          <td width="37" class="tb_subtit">&nbsp;</td>-->
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">TS1</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">TS2</div></td>
                          <!--                          <td width="37" class="tb_subtit">&nbsp;</td>-->
                          <td width="31" class="tb_subtit"> 
                            <div align="center">M1</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">Bon</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">M2</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">Rec</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">M3</div></td>
                          <td width="242" class="tb_subtit"> <div align="center">Alterado 
                              por</div></td>
                          <td width="115" class="tb_subtit"> <div align="center">Data/Hora</div></td>
                        </tr>
<!--                        <tr>
                          <td width="125" class="tb_subtit">&nbsp;</td> 
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">P</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center"> 
                              M</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">P</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37"  class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37"  class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="176" class="tb_subtit">&nbsp;</td>
                          <td width="115" class="tb_subtit">&nbsp;</td>
                        </tr>
-->                        <%
		rec_lancado="sim"
		
				Set RSprog = Server.CreateObject("ADODB.Recordset")
				SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
				RSprog.Open SQLprog, CON0
		
		check=1
			
		while not RSprog.EOF
		
				materia=RSprog("CO_Materia")
				mae=RSprog("IN_MAE")
				fil=RSprog("IN_FIL")
				in_co=RSprog("IN_CO")
				nu_peso=RSprog("NU_Peso")
				ordem=RSprog("NU_Ordem_Boletim")
				
			peso_acumula=0
			m1_ac=0
			m2_ac=0			
			m3_ac=0
			m1_exibe=0
			m2_exibe=0
			m3_exibe=0
					
			if mae=TRUE THEN
			
			check=check+1
			
					Set RS1a = Server.CreateObject("ADODB.Recordset")
					SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia_Principal='"&materia&"'"
					RS1a.Open SQL1a, CON0
					
			if RS1a.EOF then
			
						Set RS1b = Server.CreateObject("ADODB.Recordset")
						SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
						RS1b.Open SQL1b, CON0
						
						no_materia=RS1b("NO_Materia")
						
						 if check mod 2 =0 then
						  cor = "tb_fundo_linha_par" 
						 else cor ="tb_fundo_linha_impar"
						  end if
							
								Set RSnFIL = Server.CreateObject("ADODB.Recordset")
								Set RS3 = Server.CreateObject("ADODB.Recordset")
								SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' and NU_Periodo="&periodo_check
								Set RS3 = CON_N.Execute(SQL_N)
												if RS3.EOF then
								va_pt="&nbsp;"
								va_pp1="&nbsp;"
								va_pp2="&nbsp;"								
								va_t1="&nbsp;"
								va_t2="&nbsp;"							
								va_mt="&nbsp;"
								va_p1="&nbsp;"
								va_p3="&nbsp;"
								va_m1="&nbsp;"
								va_bon="&nbsp;"
								va_m2="&nbsp;"
								va_rec="&nbsp;"
								va_m3="&nbsp;"
								data_grav="&nbsp;"
								hora_grav="&nbsp;"
								usuario_grav="&nbsp;"			
						else
								va_pt=RS3("PE_Teste")
								va_pp1=RS3("PE_Prova1")
								va_pp2=RS3("PE_Prova2")								
								va_t1=RS3("VA_Teste1")
								va_t2=RS3("VA_Teste2")						
								va_mt=RS3("MD_Teste")
								va_p1=RS3("VA_Prova1")
								va_p3=RS3("VA_Prova2")
								va_m1=RS3("VA_Media1")
								va_bon=RS3("VA_Bonus")
								va_m2=RS3("VA_Media2")
								va_rec=RS3("VA_Rec")
								va_m3=RS3("VA_Media3")
								data_grav=RS3("DA_Ult_Acesso")
								hora_grav=RS3("HO_ult_Acesso")
								usuario_grav=RS3("CO_Usuario")
						end if

						
															
								
						if hora_grav="&nbsp;" then
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
											
						if data_grav="&nbsp;" then
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
						
						showapr="s"
						showprova="s"
						'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
						'data_inicio=""
						'va_faltas=""
						'		end if
						
						if usuario_grav="&nbsp;" then
						no_usuario=""
						else
								Set RS_pro = Server.CreateObject("ADODB.Recordset")
								SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
								RS_pro.Open SQL_pro, CON
							
							if RS_pro.eof then
								no_usuario= usuario_grav	
							else
								no_usuario=RS_pro("NO_Usuario")
							end if	
						end if
						%>
												<tr class="<%response.Write(cor)%>"> 
												  <td width="125"> 
													<%response.Write(no_materia)%>
												  </td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then							
													response.Write(va_t1)
													End IF							
													%>
													</div></td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_t2)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_mt)
													end if
													%>
													</div></td>
<!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													'response.Write(va_pt)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													response.Write(va_p1)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													response.Write(va_p3)
													end if
													%>
													</div></td>
												  <!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													'response.Write(va_pp)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
if isnumeric(va_m1) then
									if va_m1<val_param then
										response.Write("<font color=#F00>"& va_m1&"</font>")	
									else
										response.Write(va_m1)									
									end if									
								else	
									response.Write(va_m1)
								end if	
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_bon)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
								if isnumeric(va_m2) then
								    va_m2 = va_m2*1
									val_param = val_param*1
									if va_m2<val_param then
										response.Write("<font color=#F00>"& va_m2&"</font>")	
									else
										response.Write(va_m2)									
									end if									
								else	
									response.Write(va_m2)
								end if	
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_rec)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
								if isnumeric(va_m3) then
								    va_m3 = va_m3*1
									val_param = val_param*1
									if va_m3<val_param then
										response.Write("<font color=#F00>"& va_m3&"</font>")	
									else
										response.Write(va_m3)									
									end if									
								else	
									response.Write(va_m3)
								end if	
													end if
													%>
													</div></td>
												  <td width="242"
						> <div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then
													response.Write(no_usuario)
													end if
													%>
													</div></td>
												  <td width="115"
						> <div align="center"> 
													<%
													if showprova="s" AND showapr="s" then							
													response.Write(data_inicio)
													End if
													%></div>
												  </td>
												</tr>
					<%
			else

			
			
				 if check mod 2 =0 then
				  cor = "tb_fundo_linha_par" 
				 else cor ="tb_fundo_linha_impar"
				  end if
			
						Set RS1b = Server.CreateObject("ADODB.Recordset")
						SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
						RS1b.Open SQL1b, CON0
						
				no_materia=RS1b("NO_Materia")
					
						Set RSnFIL = Server.CreateObject("ADODB.Recordset")
						Set RS3 = Server.CreateObject("ADODB.Recordset")
						SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' and NU_Periodo="&periodo_check
						Set RS3 = CON_N.Execute(SQL_N)
										if RS3.EOF then
								va_pt="&nbsp;"
								va_pp1="&nbsp;"
								va_pp2="&nbsp;"								
								va_t1="&nbsp;"
								va_t2="&nbsp;"							
								va_mt="&nbsp;"
								va_p1="&nbsp;"
								va_p3="&nbsp;"
								va_m1="&nbsp;"
								va_bon="&nbsp;"
								va_m2="&nbsp;"
								va_rec="&nbsp;"
								va_m3="&nbsp;"
								data_grav="&nbsp;"
								hora_grav="&nbsp;"
								usuario_grav="&nbsp;"			
						else
								va_pt=RS3("PE_Teste")
								va_pp1=RS3("PE_Prova1")
								va_pp2=RS3("PE_Prova2")								
								va_t1=RS3("VA_Teste1")
								va_t2=RS3("VA_Teste2")						
								va_mt=RS3("MD_Teste")
								va_p1=RS3("VA_Prova1")
								va_p3=RS3("VA_Prova2")
								va_m1=RS3("VA_Media1")
								va_bon=RS3("VA_Bonus")
								va_m2=RS3("VA_Media2")
								va_rec=RS3("VA_Rec")
								va_m3=RS3("VA_Media3")
								data_grav=RS3("DA_Ult_Acesso")
								hora_grav=RS3("HO_ult_Acesso")
								usuario_grav=RS3("CO_Usuario")
						end if


						
				if hora_grav="&nbsp;" then
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
									
				if data_grav="&nbsp;" then
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
				
				showapr="s"
				showprova="s"
				'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
				'data_inicio=""
				'va_faltas=""
				'		end if
				
				if usuario_grav="&nbsp;" then
				no_usuario=""
				else
						Set RS_pro = Server.CreateObject("ADODB.Recordset")
						SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
						RS_pro.Open SQL_pro, CON
				
							if RS_pro.eof then
								no_usuario= usuario_grav	
							else
								no_usuario=RS_pro("NO_Usuario")
							end if	
				end if
				%>
										<tr class="<%response.Write(cor)%>"> 
										  <td width="125"> 
											<%response.Write(no_materia)%>
										  </td>
										  <td width="31"> 
											<div align="center"> 
											  <%
											if showapr="s" then							
											response.Write(va_t1)
											End IF							
											%>
											</div></td>
										  <td width="31"> 
											<div align="center"> 
											  <%
											if showapr="s" then					
											response.Write(va_t2)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showapr="s" then					
											response.Write(va_mt)
											end if
											%>
											</div></td>
<!--										  <td width="37"
				> 
											<div align="center"> 
											  <%
											if showapr="s" then					
											'response.Write(va_pt)
											end if
											%>
											</div></td>-->
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" then					
											response.Write(va_p1)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" then					
											response.Write(va_p3)
											end if
											%>
											</div></td>
										  <!--										  <td width="37"
				> 
											<div align="center"> 
											  <%
											if showprova="s" then					
											'response.Write(va_pp)
											end if
											%>
											</div></td>-->
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_m1)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_bon)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_m2)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%

											if showprova="s" AND showapr="s" then					
											response.Write(va_rec)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_m3)
											end if
											%>
											</div></td>
										  <td width="242"
				> <div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then
											response.Write(no_usuario)
											end if
											%>
											</div></td>
										  <td width="115"
				> <div align="center"> 
											<%
											if showprova="s" AND showapr="s" then							
											response.Write(data_inicio)
											End if
											%></div>
										  </td>
										</tr>
			<%
			peso_acumula=0
			acumula_m1=0
			m2_ac=0			
			m3_ac=0
			m1_exibe=0
			m2_exibe=0
			m3_exibe=0
			
			
				while not RS1a.EOF
				
						materia_fil=RS1a("CO_Materia")
					
								Set RS1b = Server.CreateObject("ADODB.Recordset")
								SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia_fil&"'"
								RS1b.Open SQL1b, CON0
								
						no_materia_fil=RS1b("NO_Materia")
						
						Set RSpa = Server.CreateObject("ADODB.Recordset")
						SQLpa= "SELECT * FROM TB_Programa_Aula where CO_Materia='"&materia_fil&"' AND CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"'"
						RSpa.Open SQLpa, CON0
												
						nu_peso_fil=RSpa("NU_Peso")	
						
						if isnull(nu_peso_fil) or nu_peso_fil="" then
							nu_peso_fil=1
						end if													
							
								Set RSnFIL = Server.CreateObject("ADODB.Recordset")
								Set RS3 = Server.CreateObject("ADODB.Recordset")
								SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia_Principal = '"& materia &"' AND CO_Materia = '"& materia_fil &"' and NU_Periodo="&periodo_check
								Set RS3 = CON_N.Execute(SQL_N)
												if RS3.EOF then
								va_pt="&nbsp;"
								va_pp1="&nbsp;"
								va_pp2="&nbsp;"								
								va_t1="&nbsp;"
								va_t2="&nbsp;"							
								va_mt="&nbsp;"
								va_p1="&nbsp;"
								va_p3="&nbsp;"
								va_m1="&nbsp;"
								va_bon="&nbsp;"
								va_m2="&nbsp;"
								va_rec="&nbsp;"
								va_m3="&nbsp;"
								data_grav="&nbsp;"
								hora_grav="&nbsp;"
								usuario_grav="&nbsp;"			
						else
								va_pt=RS3("PE_Teste")
								va_pp1=RS3("PE_Prova1")
								va_pp2=RS3("PE_Prova2")								
								va_t1=RS3("VA_Teste1")
								va_t2=RS3("VA_Teste2")						
								va_mt=RS3("MD_Teste")
								va_p1=RS3("VA_Prova1")
								va_p3=RS3("VA_Prova2")
								va_m1=RS3("VA_Media1")
								va_bon=RS3("VA_Bonus")
								va_m2=RS3("VA_Media2")
								va_rec=RS3("VA_Rec")
								va_m3=RS3("VA_Media3")
								data_grav=RS3("DA_Ult_Acesso")
								hora_grav=RS3("HO_ult_Acesso")
								usuario_grav=RS3("CO_Usuario")
						end if

						if isnull(va_m1) or va_m1="" then
							'va_m1=0
							sem_media1="s"
						else
							sem_media1="n"
							m1_ac=m1_ac+(va_m1*nu_peso_fil)															
						end if

						if isnull(va_m2) or va_m2="" then
							'va_m2=0
							sem_media2="s"
						else
							sem_media2="n"	
							m2_ac=m2_ac+(va_m2*nu_peso_fil)							
						end if
						
						if isnull(va_m3) or va_m3="" then
							'va_m3=0
							sem_media3="s"
						else
							sem_media3="n"		
							m3_ac=m3_ac+(va_m3*nu_peso_fil)					
						end if												
						
							peso_acumula=peso_acumula+nu_peso_fil
													
								
						if hora_grav="&nbsp;" then
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
											
						if data_grav="&nbsp;" then
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
						
						showapr="s"
						showprova="s"
						'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
						'data_inicio=""
						'va_faltas=""
						'		end if
						
						if usuario_grav="&nbsp;" then
						no_usuario=""
						else
								Set RS_pro = Server.CreateObject("ADODB.Recordset")
								SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
								RS_pro.Open SQL_pro, CON
						
							if RS_pro.eof then
								no_usuario= usuario_grav	
							else
								no_usuario=RS_pro("NO_Usuario")
							end if	
						end if
						%>
												
											<tr class="<%response.Write(cor)%>"> 
											  <td width="125">&nbsp;&nbsp;&nbsp;
												  <%response.Write(no_materia_fil)%>
											  </td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then							
													response.Write(va_t1)
													End IF							
													%>
													</div></td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_t2)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_mt)
													end if
													%>
													</div></td>
<!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													'response.Write(va_pt)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													response.Write(va_p1)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													response.Write(va_p3)
													end if
													%>
													</div></td>
											  <!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													'response.Write(va_pp)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m1)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_bon)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m2)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_rec)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m3)
													end if
													%>
													</div></td>
												  <td width="242"
						> <div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then
													response.Write(no_usuario)
													end if
													%>
													</div></td>
												  <td width="115"
						> <div align="center"> 
													<%
													if showprova="s" AND showapr="s" then							
													response.Write(data_inicio)
													End if
													%></div>
											  </td>
						</tr>
				<%
				RS1a.movenext
				wend
						if	sem_media1="s" then
							m1_exibe="&nbsp;"
						else
							m1_exibe=m1_ac/peso_acumula
							decimo = m1_exibe - Int(m1_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m1_exibe) + 1
									m1_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m1_exibe)
									m1_exibe=nota_arredondada					
								End If
							m1_exibe= formatNumber(m1_exibe,0)							
						end if	
							
						if	sem_media2="s" then	
							m2_exibe="&nbsp;"
						else												
							m2_exibe=m2_ac/peso_acumula
							decimo = m2_exibe - Int(m2_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m2_exibe) + 1
									m2_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m2_exibe)
									m2_exibe=nota_arredondada					
								End If
							m2_exibe= formatNumber(m2_exibe,0)							
						end if	
						
						if	sem_media3="s" then
							m3_exibe="&nbsp;"
						else							
							m3_exibe=m3_ac/peso_acumula
								decimo = m3_exibe - Int(m3_exibe)
									If decimo >= 0.5 Then
										nota_arredondada = Int(m3_exibe) + 1
										m3_exibe=nota_arredondada
									Else
										nota_arredondada = Int(m3_exibe)
										m3_exibe=nota_arredondada					
									End If
								m3_exibe= formatNumber(m3_exibe,0)									
						end if														
				
				%>
									<tr class="tb_fundo_linha_media"> 
									  <td width="125">&nbsp;&nbsp;&nbsp; M&eacute;dia</td>
										  <td width="31"> 
											<div align="center"></div></td>
										  <td width="31"> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
<!--										  <td width="37"
				> 
											<div align="center"> </div></td>-->
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
									  <!--										  <td width="37"
				> 
											<div align="center"> </div></td>-->
										  <td width="31"
				> 
											<div align="center"><%if isnumeric(m1_exibe) then
									m1_exibe = m1_exibe*1	
									val_param = val_param*1												
									if m1_exibe<val_param then
										response.Write("<font color=#F00>"& m1_exibe&"</font>")	
									else
										response.Write(m1_exibe)									
									end if									
								else	
									response.Write(m1_exibe)
								end if	%> </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center">
                              <%if isnumeric(m2_exibe) then 
									m2_exibe = m2_exibe*1	
									val_param = val_param*1	
									if m2_exibe<val_param then
										response.Write("<font color=#F00>"& m2_exibe&"</font>")	
									else
										response.Write(m2_exibe)									
									end if									
								else	
									response.Write(m2_exibe)
								end if	%>
                            </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center">
                              <%if isnumeric(m3_exibe) then 
									m3_exibe = m3_exibe*1	
									val_param = val_param*1	
									if m3_exibe<val_param then
										response.Write("<font color=#F00>"& m3_exibe&"</font>")	
									else
										response.Write(m3_exibe)									
									end if									
								else	
									response.Write(m3_exibe)
								end if	%>
                            </div></td>
										  <td width="242"
				> <div align="center"> </div></td>
										  <td width="115"
				> <div align="center"> </div>
									  </td>
						</tr>
			<%
			end if
			end if

		RSprog.MOVENEXT
		wend
		%>
								<tr valign="bottom"> 
								  <td height="20" colspan="19" 
		> <div align="right"><font class="form_corpo"> 
        
        <% if etapa=6 or etapa=7 or etapa=8 or etapa=9 then
				Response.Write("T-Teste, PR-Prova, S - Simulado, MT=(T1+T2+T3+T4)/4, MP=(PR1+S), M3=((MTx1)+(MPx2))/3. <br>Para a Disciplina Portugu�s PR2 = Reda��o e M3=((MTx1)+(MPx2)+(PR2x2))/5.")			
		else
			Response.Write("T-Teste, MT�M�dia dos Testes, PR-Prova, MP�M�dia das Provas e M-M&eacute;dia")
		End if%>        
        
        </font></div></td>
								</tr>
					  </table>
<%
elseif notaFIL="TB_NOTA_K" then
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td width="125" class="tb_subtit"> <div align="left">Disciplina</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">AV1</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">AV2</div></td>
                          <td width="31" class="tb_subtit"> <div align="center">AV3</div></td>
<!--                          <td width="37" class="tb_subtit">&nbsp;</td>-->
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">AV4</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">AV5</div></td>
                          <td width="31" align="center" class="tb_subtit">MAV</td>
                          <td width="31" align="center" class="tb_subtit">SIM</td>
                          <td width="31" align="center" class="tb_subtit">BAT</td>
                          <!--                          <td width="37" class="tb_subtit">&nbsp;</td>-->
                          <td width="31" class="tb_subtit"> 
                            <div align="center">M1</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">Bon</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">M2</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">Rec</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">M3</div></td>
                          <td width="153" class="tb_subtit"> <div align="center">Alterado 
                              por</div></td>
                          <td width="115" class="tb_subtit"> <div align="center">Data/Hora</div></td>
                        </tr>
<!--                        <tr>
                          <td width="125" class="tb_subtit">&nbsp;</td> 
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">P</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center"> 
                              M</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">P</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37"  class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37"  class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="176" class="tb_subtit">&nbsp;</td>
                          <td width="115" class="tb_subtit">&nbsp;</td>
                        </tr>
-->                        <%
		rec_lancado="sim"
		
				Set RSprog = Server.CreateObject("ADODB.Recordset")
				SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
				RSprog.Open SQLprog, CON0
		
		check=1
			
		while not RSprog.EOF
		
				materia=RSprog("CO_Materia")
				mae=RSprog("IN_MAE")
				fil=RSprog("IN_FIL")
				in_co=RSprog("IN_CO")
				nu_peso=RSprog("NU_Peso")
				ordem=RSprog("NU_Ordem_Boletim")
				
			peso_acumula=0
			m1_ac=0
			m2_ac=0			
			m3_ac=0
			m1_exibe=0
			m2_exibe=0
			m3_exibe=0
					
			if mae=TRUE THEN
			
			check=check+1
			
					Set RS1a = Server.CreateObject("ADODB.Recordset")
					SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia_Principal='"&materia&"'"
					RS1a.Open SQL1a, CON0
					
			if RS1a.EOF then
			
						Set RS1b = Server.CreateObject("ADODB.Recordset")
						SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
						RS1b.Open SQL1b, CON0
						
						no_materia=RS1b("NO_Materia")
						
						 if check mod 2 =0 then
						  cor = "tb_fundo_linha_par" 
						 else cor ="tb_fundo_linha_impar"
						  end if
							
								Set RSnFIL = Server.CreateObject("ADODB.Recordset")
								Set RS3 = Server.CreateObject("ADODB.Recordset")
								SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' and NU_Periodo="&periodo_check
								Set RS3 = CON_N.Execute(SQL_N)
						if RS3.EOF then
								va_av1="&nbsp;"
								va_av2="&nbsp;"
								va_av3="&nbsp;"								
								va_av4="&nbsp;"
								va_av5="&nbsp;"							
								va_mav="&nbsp;"
								va_sim="&nbsp;"
								va_bat="&nbsp;"
								va_m1="&nbsp;"
								va_bon="&nbsp;"
								va_m2="&nbsp;"
								va_rec="&nbsp;"
								va_m3="&nbsp;"
								data_grav="&nbsp;"
								hora_grav="&nbsp;"
								usuario_grav="&nbsp;"			
						else
						
								va_av1=RS3("VA_Av1")
								va_av2=RS3("VA_Av2")
								va_av3=RS3("VA_Av3")								
								va_av4=RS3("VA_Av4")
								va_av5=RS3("VA_Av5")						
								va_mav=RS3("VA_Mav")
								va_sim=RS3("VA_Sim")
								va_bat=RS3("VA_Bat")
								va_m1=RS3("VA_Media1")
								va_bon=RS3("VA_Bonus")
								va_m2=RS3("VA_Media2")
								va_rec=RS3("VA_Rec")
								va_m3=RS3("VA_Media3")
								data_grav=RS3("DA_Ult_Acesso")
								hora_grav=RS3("HO_ult_Acesso")
								usuario_grav=RS3("CO_Usuario")
						end if

						
															
								
						if hora_grav="&nbsp;" or hora_grav="" or isnull(hora_grav) then
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
											
						if data_grav="&nbsp;" or data_grav="" or isnull(data_grav) then
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
						
						showapr="s"
						showprova="s"
						'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
						'data_inicio=""
						'va_faltas=""
						'		end if
						
						if usuario_grav="&nbsp;" then
						no_usuario=""
						else
								Set RS_pro = Server.CreateObject("ADODB.Recordset")
								SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
								RS_pro.Open SQL_pro, CON
							
							if RS_pro.eof then
								no_usuario= usuario_grav	
							else
								no_usuario=RS_pro("NO_Usuario")
							end if	
						end if
						%>
												<tr class="<%response.Write(cor)%>"> 
												  <td width="125"> 
													<%response.Write(no_materia)%>
												  </td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then							
													response.Write(va_av1)
													End IF							
													%>
													</div></td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_av2)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_av3)
													end if
													%>
													</div></td>
<!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													'response.Write(va_pt)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_av4)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_av5)
													end if
													%>
													</div></td>
												  <td width="31" align="center"
						><%if showapr="s" then					
													response.Write(va_mav)
													end if
													%></td>
												  <td width="31" align="center"
						><%if showprova="s" then					
													response.Write(va_sim)
													end if
													%></td>
												  <td width="31" align="center"
						><%if showprova="s" then					
													response.Write(va_bat)
													end if
													%><strong></strong></td>
												  <!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													'response.Write(va_pp)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%

													if showprova="s" AND showapr="s" then					
														if isnumeric(va_m1) then
															if va_m1<val_param then
																response.Write("<font color=#F00>"& va_m1&"</font>")	
															else
																response.Write(va_m1)									
															end if									
														else	
															response.Write(va_m1)
														end if	
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_bon)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
								if isnumeric(va_m2) then
								    va_m2 = va_m2*1
									val_param = val_param*1
									if va_m2<val_param then
										response.Write("<font color=#F00>"& va_m2&"</font>")	
									else
										response.Write(va_m2)									
									end if									
								else	
									response.Write(va_m2)
								end if	
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_rec)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
								if isnumeric(va_m3) then
								    va_m3 = va_m3*1
									val_param = val_param*1
									if va_m3<val_param then
										response.Write("<font color=#F00>"& va_m3&"</font>")	
									else
										response.Write(va_m3)									
									end if									
								else	
									response.Write(va_m3)
								end if	
													end if
													%>
													</div></td>
												  <td width="153"
						> <div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then
													response.Write(no_usuario)
													end if
													%>
													</div></td>
												  <td width="115"
						> <div align="center"> 
													<%
													if showprova="s" AND showapr="s" then							
													response.Write(data_inicio)
													End if
													%></div>
												  </td>
												</tr>
					<%
			else

			
			
				 if check mod 2 =0 then
				  cor = "tb_fundo_linha_par" 
				 else cor ="tb_fundo_linha_impar"
				  end if
			
						Set RS1b = Server.CreateObject("ADODB.Recordset")
						SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
						RS1b.Open SQL1b, CON0
						
				no_materia=RS1b("NO_Materia")
					
						Set RSnFIL = Server.CreateObject("ADODB.Recordset")
						Set RS3 = Server.CreateObject("ADODB.Recordset")
						SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' and NU_Periodo="&periodo_check
						Set RS3 = CON_N.Execute(SQL_N)
						if RS3.EOF then
								va_av1="&nbsp;"
								va_av2="&nbsp;"
								va_av3="&nbsp;"								
								va_av4="&nbsp;"
								va_av5="&nbsp;"							
								va_mav="&nbsp;"
								va_sim="&nbsp;"
								va_bat="&nbsp;"
								va_m1="&nbsp;"
								va_bon="&nbsp;"
								va_m2="&nbsp;"
								va_rec="&nbsp;"
								va_m3="&nbsp;"
								data_grav="&nbsp;"
								hora_grav="&nbsp;"
								usuario_grav="&nbsp;"			
						else
						
								va_av1=RS3("VA_Av1")
								va_av2=RS3("VA_Av2")
								va_av3=RS3("VA_Av3")								
								va_av4=RS3("VA_Av4")
								va_av5=RS3("VA_Av5")						
								va_mav=RS3("VA_Mav")
								va_sim=RS3("VA_Sim")
								va_bat=RS3("VA_Bat")
								va_m1=RS3("VA_Media1")
								va_bon=RS3("VA_Bonus")
								va_m2=RS3("VA_Media2")
								va_rec=RS3("VA_Rec")
								va_m3=RS3("VA_Media3")
								data_grav=RS3("DA_Ult_Acesso")
								hora_grav=RS3("HO_ult_Acesso")
								usuario_grav=RS3("CO_Usuario")
						end if


						
				if hora_grav="&nbsp;" or hora_grav="" or isnull(hora_grav) then
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
									
				if data_grav="&nbsp;" or data_grav="" or isnull(data_grav)  then
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
				
				showapr="s"
				showprova="s"
				'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
				'data_inicio=""
				'va_faltas=""
				'		end if
				
				if usuario_grav="&nbsp;" then
				no_usuario=""
				else
						Set RS_pro = Server.CreateObject("ADODB.Recordset")
						SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
						RS_pro.Open SQL_pro, CON
				
							if RS_pro.eof then
								no_usuario= usuario_grav	
							else
								no_usuario=RS_pro("NO_Usuario")
							end if	
				end if
				%>
										<tr class="<%response.Write(cor)%>"> 
										  <td width="125"> 
											<%response.Write(no_materia)%>
										  </td>
										  <td width="31"> 
											<div align="center"> 
											  <%
											if showapr="s" then							
											response.Write(va_av1)
											End IF							
											%>
											</div></td>
										  <td width="31"> 
											<div align="center"> 
											  <%
											if showapr="s" then					
											response.Write(va_av2)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showapr="s" then					
											response.Write(va_av3)
											end if
											%>
											</div></td>
<!--										  <td width="37"
				> 
											<div align="center"> 
											  <%
											if showapr="s" then					
											'response.Write(va_pt)
											end if
											%>
											</div></td>-->
										  <td width="31"
				> 

											<div align="center"> 
											  <%
											if showapr="s" then					
											response.Write(va_av4)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showapr="s" then					
											response.Write(va_av5)
											end if
											%>
											</div></td>
										  <td align="center"
						><%if showapr="s" then					
													response.Write(va_mav)
													end if
													%></td>
										  <td align="center"
						><%if showprova="s" then					
													response.Write(va_sim)
													end if
													%></td>
										  <td align="center"
						><%if showprova="s" then					
													response.Write(va_bat)
													end if
													%>
										    <strong></strong></td>
										  <!--										  <td width="37"
				> 
											<div align="center"> 
											  <%
											if showprova="s" then					
											'response.Write(va_pp)
											end if
											%>
											</div></td>-->
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_m1)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_bon)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_m2)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%

											if showprova="s" AND showapr="s" then					
											response.Write(va_rec)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_m3)
											end if
											%>
											</div></td>
										  <td width="153"
				> <div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then
											response.Write(no_usuario)
											end if
											%>
											</div></td>
										  <td width="115"
				> <div align="center"> 
											<%
											if showprova="s" AND showapr="s" then							
											response.Write(data_inicio)
											End if
											%></div>
										  </td>
										</tr>
			<%
			peso_acumula=0
			acumula_m1=0
			m2_ac=0			
			m3_ac=0
			m1_exibe=0
			m2_exibe=0
			m3_exibe=0
			
			
				while not RS1a.EOF
				
						materia_fil=RS1a("CO_Materia")
					
								Set RS1b = Server.CreateObject("ADODB.Recordset")
								SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia_fil&"'"
								RS1b.Open SQL1b, CON0
								
						no_materia_fil=RS1b("NO_Materia")
						
						Set RSpa = Server.CreateObject("ADODB.Recordset")
						SQLpa= "SELECT * FROM TB_Programa_Aula where CO_Materia='"&materia_fil&"' AND CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"'"
						RSpa.Open SQLpa, CON0
												
						nu_peso_fil=RSpa("NU_Peso")		
						if isnull(nu_peso_fil) or nu_peso_fil="" then
							nu_peso_fil=1
						end if												
							
								Set RSnFIL = Server.CreateObject("ADODB.Recordset")
								Set RS3 = Server.CreateObject("ADODB.Recordset")
								SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia_Principal = '"& materia &"' AND CO_Materia = '"& materia_fil &"' and NU_Periodo="&periodo_check
								Set RS3 = CON_N.Execute(SQL_N)
						if RS3.EOF then
								va_av1="&nbsp;"
								va_av2="&nbsp;"
								va_av3="&nbsp;"								
								va_av4="&nbsp;"
								va_av5="&nbsp;"							
								va_mav="&nbsp;"
								va_sim="&nbsp;"
								va_bat="&nbsp;"
								va_m1="&nbsp;"
								va_bon="&nbsp;"
								va_m2="&nbsp;"
								va_rec="&nbsp;"
								va_m3="&nbsp;"
								data_grav="&nbsp;"
								hora_grav="&nbsp;"
								usuario_grav="&nbsp;"			
						else
						
								va_av1=RS3("VA_Av1")
								va_av2=RS3("VA_Av2")
								va_av3=RS3("VA_Av3")								
								va_av4=RS3("VA_Av4")
								va_av5=RS3("VA_Av5")						
								va_mav=RS3("VA_Mav")
								va_sim=RS3("VA_Sim")
								va_bat=RS3("VA_Bat")
								va_m1=RS3("VA_Media1")
								va_bon=RS3("VA_Bonus")
								va_m2=RS3("VA_Media2")
								va_rec=RS3("VA_Rec")
								va_m3=RS3("VA_Media3")
								data_grav=RS3("DA_Ult_Acesso")
								hora_grav=RS3("HO_ult_Acesso")
								usuario_grav=RS3("CO_Usuario")
						end if
						
						if isnull(va_m1) or va_m1="" or va_m1="&nbsp;" then
							'va_m1=0
							sem_media1="s"
						else
							sem_media1="n"
							m1_ac=m1_ac+(va_m1*nu_peso_fil)															
						end if

						if isnull(va_m2) or va_m2="" or va_m2="&nbsp;" then
							'va_m2=0
							sem_media2="s"
						else
							sem_media2="n"	
							m2_ac=m2_ac+(va_m2*nu_peso_fil)							
						end if
						
						if isnull(va_m3) or va_m3="" or va_m3="&nbsp;" then
							'va_m3=0
							sem_media3="s"
						else
							sem_media3="n"		
							m3_ac=m3_ac+(va_m3*nu_peso_fil)					
						end if												
						
							'peso_acumula=peso_acumula+nu_peso_fil
						peso_acumula=nu_peso_fil								
													
								
						if hora_grav="&nbsp;" or hora_grav="" or isnull(hora_grav)  then
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
											
						if data_grav="&nbsp;" or data_grav="" or isnull(data_grav) then
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
						
						showapr="s"
						showprova="s"
						'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
						'data_inicio=""
						'va_faltas=""
						'		end if
						
						if usuario_grav="&nbsp;" then
						no_usuario=""
						else
								Set RS_pro = Server.CreateObject("ADODB.Recordset")
								SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
								RS_pro.Open SQL_pro, CON
						
							if RS_pro.eof then
								no_usuario= usuario_grav	
							else
								no_usuario=RS_pro("NO_Usuario")
							end if	
						end if
						%>
												
											<tr class="<%response.Write(cor)%>"> 
											  <td width="125">&nbsp;&nbsp;&nbsp;
												  <%response.Write(no_materia_fil)%>
											  </td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then							
													response.Write(va_av1)
													End IF							
													%>
													</div></td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_av2)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_av3)
													end if
													%>
													</div></td>
<!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													'response.Write(va_pt)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_av4)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_av5)
													end if
													%>
													</div></td>
												  <td align="center"
						><%if showapr="s" then					
													response.Write(va_mav)
													end if
													%></td>
												  <td align="center"
						><%if showprova="s" then					
													response.Write(va_sim)
													end if
													%></td>
												  <td align="center"
						><%if showprova="s" then					
													response.Write(va_bat)
													end if
													%></td>
											  <!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													'response.Write(va_pp)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m1)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_bon)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m2)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_rec)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m3)
													end if
													%>
													</div></td>
												  <td width="153"
						> <div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then
													response.Write(no_usuario)
													end if
													%>
													</div></td>
												  <td width="115"
						> <div align="center"> 
													<%
													if showprova="s" AND showapr="s" then							
													response.Write(data_inicio)
													End if
													%></div>
											  </td>
						</tr>
				<%
				RS1a.movenext
				wend
						if	sem_media1="s" then
							m1_exibe="&nbsp;"
						else
							m1_exibe=m1_ac/peso_acumula
							decimo = m1_exibe - Int(m1_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m1_exibe) + 1
									m1_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m1_exibe)
									m1_exibe=nota_arredondada					
								End If
							if m1_exibe>100 then
								m1_exibe=100
							end if
							m1_exibe= formatNumber(m1_exibe,0)							
						end if	
							
						if	sem_media2="s" then	
							m2_exibe="&nbsp;"
						else												
							m2_exibe=m2_ac/peso_acumula
							decimo = m2_exibe - Int(m2_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m2_exibe) + 1
									m2_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m2_exibe)
									m2_exibe=nota_arredondada					
								End If
							if m2_exibe>100 then
								m2_exibe=100
							end if								
							m2_exibe= formatNumber(m2_exibe,0)							
						end if	
						
						if	sem_media3="s" then
							m3_exibe="&nbsp;"
						else							
							m3_exibe=m3_ac/peso_acumula
								decimo = m3_exibe - Int(m3_exibe)
									If decimo >= 0.5 Then
										nota_arredondada = Int(m3_exibe) + 1
										m3_exibe=nota_arredondada
									Else
										nota_arredondada = Int(m3_exibe)
										m3_exibe=nota_arredondada					
									End If
								if m3_exibe>100 then
									m3_exibe=100
								end if										
								m3_exibe= formatNumber(m3_exibe,0)									
						end if														
				
				%>
									<tr class="tb_fundo_linha_media"> 
									  <td width="125">&nbsp;&nbsp;&nbsp; M&eacute;dia</td>
										  <td width="31"> 
											<div align="center"></div></td>
										  <td width="31"> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
<!--										  <td width="37"
				> 
											<div align="center"> </div></td>-->
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				>&nbsp;</td>
										  <td width="31"
				>&nbsp;</td>
										  <td width="31"
				>&nbsp;</td>
									  <!--										  <td width="37"
				> 
											<div align="center"> </div></td>-->
										  <td width="31"
				> 
											<div align="center"><%if isnumeric(m1_exibe) then 
									m1_exibe = m1_exibe*1	
									val_param = val_param*1	
									if m1_exibe<val_param then
										response.Write("<font color=#F00>"& m1_exibe&"</font>")	
									else
										response.Write(m1_exibe)									
									end if									
								else	
									response.Write(m1_exibe)
								end if	%> </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center">
                              <%if isnumeric(m2_exibe) then 
									m2_exibe = m2_exibe*1	
									val_param = val_param*1	
									if m2_exibe<val_param then
										response.Write("<font color=#F00>"& m2_exibe&"</font>")	
									else
										response.Write(m2_exibe)									
									end if									
								else	
									response.Write(m2_exibe)
								end if	%>
                            </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center">
                              <%if isnumeric(m3_exibe) then 
									m3_exibe = m3_exibe*1	
									val_param = val_param*1	
									if m3_exibe<val_param then
										response.Write("<font color=#F00>"& m3_exibe&"</font>")	
									else
										response.Write(m3_exibe)									
									end if									
								else	
									response.Write(m3_exibe)
								end if	%>
                            </div></td>
										  <td width="153"
				> <div align="center"> </div></td>
										  <td width="115"
				> <div align="center"> </div>
									  </td>
						</tr>
			<%
			end if
			end if

		RSprog.MOVENEXT
		wend
		%>
								<tr valign="bottom"> 
								  <td height="20" colspan="22" 
		> <div align="right"><font class="form_corpo"> 
        
        <% if etapa=6 or etapa=7 or etapa=8 or etapa=9 then
				Response.Write("T-Teste, PR-Prova, S - Simulado, MT=(T1+T2+T3+T4)/4, MP=(PR1+S), M3=((MTx1)+(MPx2))/3. <br>Para a Disciplina Portugu�s PR2 = Reda��o e M3=((MTx1)+(MPx2)+(PR2x2))/5.")			
		else
			Response.Write("AV-Avalia&ccedil;&otilde;es, MAV--M&eacute;dia das Avalia&ccedil;&otilde;es, SIM-Simulado, BAT-Bonus Atualidade  e M-M&eacute;dia")
		End if%>        
        
        </font></div></td>
								</tr>
					  </table>
<%
elseif notaFIL="TB_NOTA_L" then
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td width="125" class="tb_subtit"> <div align="left">Disciplina</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">T1</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">T2</div></td>
                          <td width="31" class="tb_subtit"> <div align="center">T3</div></td>
<!--                          <td width="37" class="tb_subtit">&nbsp;</td>-->
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">T4</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">MT</div></td>
                          <td width="31" align="center" class="tb_subtit">P1</td>
                          <td width="31" align="center" class="tb_subtit">P2</td>
                          <td width="31" align="center" class="tb_subtit">MP</td>
                          <td width="31" align="center" class="tb_subtit">SIM</td>
                          <!--                          <td width="37" class="tb_subtit">&nbsp;</td>-->
                          <td width="31" class="tb_subtit"> 
                            <div align="center">M1</div></td>
                          <td width="31" align="center" class="tb_subtit">BAT</td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">Bon</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">M2</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">Rec</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">M3</div></td>
                          <td width="153" class="tb_subtit"> <div align="center">Alterado 
                              por</div></td>
                          <td width="115" class="tb_subtit"> <div align="center">Data/Hora</div></td>
                        </tr>
<!--                        <tr>
                          <td width="125" class="tb_subtit">&nbsp;</td> 
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">P</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center"> 
                              M</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">P</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37"  class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37"  class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="176" class="tb_subtit">&nbsp;</td>
                          <td width="115" class="tb_subtit">&nbsp;</td>
                        </tr>
-->                        <%
		rec_lancado="sim"
		
				Set RSprog = Server.CreateObject("ADODB.Recordset")
				SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
				RSprog.Open SQLprog, CON0
		
		check=1
			
		while not RSprog.EOF
		
				materia=RSprog("CO_Materia")
				mae=RSprog("IN_MAE")
				fil=RSprog("IN_FIL")
				in_co=RSprog("IN_CO")
				nu_peso=RSprog("NU_Peso")
				ordem=RSprog("NU_Ordem_Boletim")
				
			peso_acumula=0
			m1_ac=0
			m2_ac=0			
			m3_ac=0
			m1_exibe=0
			m2_exibe=0
			m3_exibe=0
					
			if mae=TRUE THEN
			
			check=check+1
			
					Set RS1a = Server.CreateObject("ADODB.Recordset")
					SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia_Principal='"&materia&"'"
					RS1a.Open SQL1a, CON0
					
			if RS1a.EOF then
			
						Set RS1b = Server.CreateObject("ADODB.Recordset")
						SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
						RS1b.Open SQL1b, CON0
						
						no_materia=RS1b("NO_Materia")
						
						 if check mod 2 =0 then
						  cor = "tb_fundo_linha_par" 
						 else cor ="tb_fundo_linha_impar"
						  end if
							
								Set RSnFIL = Server.CreateObject("ADODB.Recordset")
								Set RS3 = Server.CreateObject("ADODB.Recordset")
								SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' and NU_Periodo="&periodo_check
								Set RS3 = CON_N.Execute(SQL_N)
						if RS3.EOF then
								va_t1="&nbsp;"
								va_t2="&nbsp;"
								va_t3="&nbsp;"								
								va_t4="&nbsp;"
								va_mt="&nbsp;"	
								va_p1="&nbsp;"	
								va_p2="&nbsp;"	
								va_mp="&nbsp;"	
								va_sim="&nbsp;"
								va_m1="&nbsp;"								
								va_bat="&nbsp;"
								va_bon="&nbsp;"
								va_m2="&nbsp;"
								va_rec="&nbsp;"
								va_m3="&nbsp;"
								data_grav="&nbsp;"
								hora_grav="&nbsp;"
								usuario_grav="&nbsp;"			
						else
						
								va_t1=RS3("VA_Teste1")
								va_t2=RS3("VA_Teste2")
								va_t3=RS3("VA_Teste3")								
								va_t4=RS3("VA_Teste4")
								va_mt=RS3("MD_Teste")	
								va_p1=RS3("VA_Prova1")	
								va_p2=RS3("VA_Prova2")
								va_mp=RS3("MD_Prova")	
								va_sim=RS3("VA_Sim")
								va_m1=RS3("VA_Media1")					
								va_bat=RS3("VA_Bat")
								va_bon=RS3("VA_Bonus")
								va_m2=RS3("VA_Media2")
								va_rec=RS3("VA_Rec")
								va_m3=RS3("VA_Media3")
								data_grav=RS3("DA_Ult_Acesso")
								hora_grav=RS3("HO_ult_Acesso")
								usuario_grav=RS3("CO_Usuario")
						end if

						
															
								
						if hora_grav="&nbsp;" or hora_grav="" or isnull(hora_grav) then
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
											
						if data_grav="&nbsp;" or data_grav="" or isnull(data_grav) then
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
						
						showapr="s"
						showprova="s"
						'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
						'data_inicio=""
						'va_faltas=""
						'		end if
						
						if usuario_grav="&nbsp;" then
						no_usuario=""
						else
								Set RS_pro = Server.CreateObject("ADODB.Recordset")
								SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
								RS_pro.Open SQL_pro, CON
							
							if RS_pro.eof then
								no_usuario= usuario_grav	
							else
								no_usuario=RS_pro("NO_Usuario")
							end if	
						end if
						%>
												<tr class="<%response.Write(cor)%>"> 
												  <td width="125"> 
													<%response.Write(no_materia)%>
												  </td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then							
													response.Write(va_t1)
													End IF							
													%>
													</div></td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_t2)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_t3)
													end if
													%>
													</div></td>
<!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													'response.Write(va_pt)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_t4)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_mt)
													end if
													%>
													</div></td>
												  <td width="31" align="center"
						><%if showprova="s" then					
													response.Write(va_p1)
													end if
													%></td>
												  <td width="31" align="center"
						><%if showprova="s" then					
													response.Write(va_p2)
													end if
													%><strong></strong></td>
												  <td width="31" align="center"
						><%if showapr="s" then					
													response.Write(va_mp)
													end if
													%></td>
												  <td width="31" align="center"
						><%if showprova="s" then					
													response.Write(va_sim)
													end if
													%></td>
												  <!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													'response.Write(va_pp)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%

													if showprova="s" AND showapr="s" then					
														if isnumeric(va_m1) then
															if va_m1<val_param then
																response.Write("<font color=#F00>"& va_m1&"</font>")	
															else
																response.Write(va_m1)									
															end if									
														else	
															response.Write(va_m1)
														end if	
													end if
													%>
													</div></td>
												  <td width="31" align="center"
						><%if showprova="s" then					
													response.Write(va_bat)
													end if
													%>
												    <strong></strong></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_bon)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
								if isnumeric(va_m2) then
								    va_m2 = va_m2*1
									val_param = val_param*1
									if va_m2<val_param then
										response.Write("<font color=#F00>"& va_m2&"</font>")	
									else
										response.Write(va_m2)									
									end if									
								else	
									response.Write(va_m2)
								end if	
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_rec)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
								if isnumeric(va_m3) then
								    va_m3 = va_m3*1
									val_param = val_param*1
									if va_m3<val_param then
										response.Write("<font color=#F00>"& va_m3&"</font>")	
									else
										response.Write(va_m3)									
									end if									
								else	
									response.Write(va_m3)
								end if	
													end if
													%>
													</div></td>
												  <td width="153"
						> <div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then
													response.Write(no_usuario)
													end if
													%>
													</div></td>
												  <td width="115"
						> <div align="center"> 
													<%
													if showprova="s" AND showapr="s" then							
													response.Write(data_inicio)
													End if
													%></div>
												  </td>
												</tr>
					<%
			else

			
			
				 if check mod 2 =0 then
				  cor = "tb_fundo_linha_par" 
				 else cor ="tb_fundo_linha_impar"
				  end if
			
						Set RS1b = Server.CreateObject("ADODB.Recordset")
						SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
						RS1b.Open SQL1b, CON0
						
				no_materia=RS1b("NO_Materia")
					
						Set RSnFIL = Server.CreateObject("ADODB.Recordset")
						Set RS3 = Server.CreateObject("ADODB.Recordset")
						SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' and NU_Periodo="&periodo_check
						Set RS3 = CON_N.Execute(SQL_N)
						if RS3.EOF then
								va_t1="&nbsp;"
								va_t2="&nbsp;"
								va_t3="&nbsp;"								
								va_t4="&nbsp;"
								va_mt="&nbsp;"	
								va_p1="&nbsp;"	
								va_p2="&nbsp;"	
								va_mp="&nbsp;"	
								va_sim="&nbsp;"
								va_m1="&nbsp;"								
								va_bat="&nbsp;"
								va_bon="&nbsp;"
								va_m2="&nbsp;"
								va_rec="&nbsp;"
								va_m3="&nbsp;"
								data_grav="&nbsp;"
								hora_grav="&nbsp;"
								usuario_grav="&nbsp;"			
						else
						
								va_t1=RS3("VA_Teste1")
								va_t2=RS3("VA_Teste2")
								va_t3=RS3("VA_Teste3")								
								va_t4=RS3("VA_Teste4")
								va_mt=RS3("MD_Teste")	
								va_p1=RS3("VA_Prova1")	
								va_p2=RS3("VA_Prova2")
								va_mp=RS3("MD_Prova")	
								va_sim=RS3("VA_Sim")
								va_m1=RS3("VA_Media1")					
								va_bat=RS3("VA_Bat")
								va_bon=RS3("VA_Bonus")
								va_m2=RS3("VA_Media2")
								va_rec=RS3("VA_Rec")
								va_m3=RS3("VA_Media3")
								data_grav=RS3("DA_Ult_Acesso")
								hora_grav=RS3("HO_ult_Acesso")
								usuario_grav=RS3("CO_Usuario")
						end if


						
				if hora_grav="&nbsp;" or hora_grav="" or isnull(hora_grav) then
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
									
				if data_grav="&nbsp;" or data_grav="" or isnull(data_grav)  then
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
				
				showapr="s"
				showprova="s"
				'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
				'data_inicio=""
				'va_faltas=""
				'		end if
				
				if usuario_grav="&nbsp;" then
				no_usuario=""
				else
						Set RS_pro = Server.CreateObject("ADODB.Recordset")
						SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
						RS_pro.Open SQL_pro, CON
				
							if RS_pro.eof then
								no_usuario= usuario_grav	
							else
								no_usuario=RS_pro("NO_Usuario")
							end if	
				end if
				%>
										<tr class="<%response.Write(cor)%>"> 
										  <td width="125"> 
											<%response.Write(no_materia)%>
										  </td>
										  <td width="31"> 
											<div align="center">
											  <%
											if showapr="s" then							
											response.Write(va_t1)
											End IF							
											%>
											</div></td>
										  <td width="31"> 
											<div align="center"> 
											  <%
											if showapr="s" then					
											response.Write(va_t2)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showapr="s" then					
											response.Write(va_t3)
											end if
											%>
											</div></td>
<!--										  <td width="37"
				> 
											<div align="center"> 
											  <%
											if showapr="s" then					
											'response.Write(va_pt)
											end if
											%>
											</div></td>-->
										  <td width="31"
				> 

											<div align="center"> 
											  <%
											if showapr="s" then					
											response.Write(va_t4)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showapr="s" then					
											response.Write(va_mt)
											end if
											%>
											</div></td>
										  <td align="center"
						><%if showprova="s" then					
													response.Write(va_p1)
													end if
													%></td>
										  <td align="center"
						><%if showprova="s" then					
													response.Write(va_p1)
													end if
													%>
										    <strong></strong></td>
										  <td align="center"
						><%if showapr="s" then					
													response.Write(va_mp)
													end if
													%></td>
										  <td align="center"
						><%if showprova="s" then					
													response.Write(va_sim)
													end if
													%></td>
										  <!--										  <td width="37"
				> 
											<div align="center"> 
											  <%
											if showprova="s" then					
											'response.Write(va_pp)
											end if
											%>
											</div></td>-->
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_m1)
											end if
											%>
											</div></td>
										  <td align="center"
						><%if showprova="s" then					
													response.Write(va_bat)
													end if
													%>
										    <strong></strong></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_bon)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_m2)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%

											if showprova="s" AND showapr="s" then					
											response.Write(va_rec)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_m3)
											end if
											%>
											</div></td>
										  <td width="153"
				> <div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then
											response.Write(no_usuario)
											end if
											%>
											</div></td>
										  <td width="115"
				> <div align="center"> 
											<%
											if showprova="s" AND showapr="s" then							
											response.Write(data_inicio)
											End if
											%></div>
										  </td>
										</tr>
			<%
			peso_acumula=0
			acumula_m1=0
			m2_ac=0			
			m3_ac=0
			m1_exibe=0
			m2_exibe=0
			m3_exibe=0
			
			
				while not RS1a.EOF
				
						materia_fil=RS1a("CO_Materia")
					
								Set RS1b = Server.CreateObject("ADODB.Recordset")
								SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia_fil&"'"
								RS1b.Open SQL1b, CON0
								
						no_materia_fil=RS1b("NO_Materia")
						
						Set RSpa = Server.CreateObject("ADODB.Recordset")
						SQLpa= "SELECT * FROM TB_Programa_Aula where CO_Materia='"&materia_fil&"' AND CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"'"
						RSpa.Open SQLpa, CON0
												
						nu_peso_fil=RSpa("NU_Peso")		
						if isnull(nu_peso_fil) or nu_peso_fil="" then
							nu_peso_fil=1
						end if												
							
								Set RSnFIL = Server.CreateObject("ADODB.Recordset")
								Set RS3 = Server.CreateObject("ADODB.Recordset")
								SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia_Principal = '"& materia &"' AND CO_Materia = '"& materia_fil &"' and NU_Periodo="&periodo_check
								Set RS3 = CON_N.Execute(SQL_N)
						if RS3.EOF then
								va_t1="&nbsp;"
								va_t2="&nbsp;"
								va_t3="&nbsp;"								
								va_t4="&nbsp;"
								va_mt="&nbsp;"	
								va_p1="&nbsp;"	
								va_p2="&nbsp;"	
								va_mp="&nbsp;"	
								va_sim="&nbsp;"
								va_m1="&nbsp;"								
								va_bat="&nbsp;"
								va_bon="&nbsp;"
								va_m2="&nbsp;"
								va_rec="&nbsp;"
								va_m3="&nbsp;"
								data_grav="&nbsp;"
								hora_grav="&nbsp;"
								usuario_grav="&nbsp;"			
						else
						
								va_t1=RS3("VA_Teste1")
								va_t2=RS3("VA_Teste2")
								va_t3=RS3("VA_Teste3")								
								va_t4=RS3("VA_Teste4")
								va_mt=RS3("MD_Teste")	
								va_p1=RS3("VA_Prova1")	
								va_p2=RS3("VA_Prova2")
								va_mp=RS3("MD_Prova")	
								va_sim=RS3("VA_Sim")
								va_m1=RS3("VA_Media1")					
								va_bat=RS3("VA_Bat")
								va_bon=RS3("VA_Bonus")
								va_m2=RS3("VA_Media2")
								va_rec=RS3("VA_Rec")
								va_m3=RS3("VA_Media3")
								data_grav=RS3("DA_Ult_Acesso")
								hora_grav=RS3("HO_ult_Acesso")
								usuario_grav=RS3("CO_Usuario")
						end if
						
						if isnull(va_m1) or va_m1="" or va_m1="&nbsp;" then
							'va_m1=0
							sem_media1="s"
						else
							sem_media1="n"
							m1_ac=m1_ac+(va_m1*nu_peso_fil)															
						end if

						if isnull(va_m2) or va_m2="" or va_m2="&nbsp;" then
							'va_m2=0
							sem_media2="s"
						else
							sem_media2="n"	
							m2_ac=m2_ac+(va_m2*nu_peso_fil)							
						end if
						
						if isnull(va_m3) or va_m3="" or va_m3="&nbsp;" then
							'va_m3=0
							sem_media3="s"
						else
							sem_media3="n"		
							m3_ac=m3_ac+(va_m3*nu_peso_fil)					
						end if												
						
							peso_acumula=peso_acumula+nu_peso_fil
						'peso_acumula=nu_peso_fil								
													
								
						if hora_grav="&nbsp;" or hora_grav="" or isnull(hora_grav)  then
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
											
						if data_grav="&nbsp;" or data_grav="" or isnull(data_grav) then
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
						
						showapr="s"
						showprova="s"
						'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
						'data_inicio=""
						'va_faltas=""
						'		end if
						
						if usuario_grav="&nbsp;" then
						no_usuario=""
						else
								Set RS_pro = Server.CreateObject("ADODB.Recordset")
								SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
								RS_pro.Open SQL_pro, CON
						
							if RS_pro.eof then
								no_usuario= usuario_grav	
							else
								no_usuario=RS_pro("NO_Usuario")
							end if	
						end if
						%>
												
											<tr class="<%response.Write(cor)%>"> 
											  <td width="125">&nbsp;&nbsp;&nbsp;
												  <%response.Write(no_materia_fil)%>
											  </td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then							
													response.Write(va_t1)
													End IF							
													%>
													</div></td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_t2)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_t3)
													end if
													%>
													</div></td>
<!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													'response.Write(va_pt)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_t4)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_mt)
													end if
													%>
													</div></td>
												  <td align="center"
						><%if showprova="s" then					
													response.Write(va_p1)
													end if
													%></td>
												  <td align="center"
						><strong>
												    <%if showprova="s" then					
													response.Write(va_p1)
													end if
													%>
												  </strong></td>
												  <td align="center"
						><%if showapr="s" then					
													response.Write(va_mp)
													end if
													%></td>
												  <td align="center"
						><%if showprova="s" then					
													response.Write(va_sim)
													end if
													%></td>
											  <!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													'response.Write(va_pp)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m1)
													end if
													%>
													</div></td>
												  <td align="center"
						><%if showprova="s" then					
													response.Write(va_bat)
													end if
													%></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_bon)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m2)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_rec)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m3)
													end if
													%>
													</div></td>
												  <td width="153"
						> <div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then
													response.Write(no_usuario)
													end if
													%>
													</div></td>
												  <td width="115"
						> <div align="center"> 
													<%
													if showprova="s" AND showapr="s" then							
													response.Write(data_inicio)
													End if
													%></div>
											  </td>
						</tr>
				<%
				RS1a.movenext
				wend
						if	sem_media1="s" then
							m1_exibe="&nbsp;"
						else
							m1_exibe=m1_ac/peso_acumula
							decimo = m1_exibe - Int(m1_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m1_exibe) + 1
									m1_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m1_exibe)
									m1_exibe=nota_arredondada					
								End If
							if m1_exibe>100 then
								m1_exibe=100
							end if
							m1_exibe= formatNumber(m1_exibe,0)							
						end if	
							
						if	sem_media2="s" then	
							m2_exibe="&nbsp;"
						else												
							m2_exibe=m2_ac/peso_acumula
							decimo = m2_exibe - Int(m2_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m2_exibe) + 1
									m2_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m2_exibe)
									m2_exibe=nota_arredondada					
								End If
							if m2_exibe>100 then
								m2_exibe=100
							end if								
							m2_exibe= formatNumber(m2_exibe,0)							
						end if	
						
						if	sem_media3="s" then
							m3_exibe="&nbsp;"
						else							
							m3_exibe=m3_ac/peso_acumula
								decimo = m3_exibe - Int(m3_exibe)
									If decimo >= 0.5 Then
										nota_arredondada = Int(m3_exibe) + 1
										m3_exibe=nota_arredondada
									Else
										nota_arredondada = Int(m3_exibe)
										m3_exibe=nota_arredondada					
									End If
								if m3_exibe>100 then
									m3_exibe=100
								end if										
								m3_exibe= formatNumber(m3_exibe,0)									
						end if														
				
				%>
									<tr class="tb_fundo_linha_media"> 
									  <td width="125">&nbsp;&nbsp;&nbsp; M&eacute;dia</td>
										  <td width="31"> 
											<div align="center"></div></td>
										  <td width="31"> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
<!--										  <td width="37"
				> 
											<div align="center"> </div></td>-->
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				>&nbsp;</td>
										  <td width="31"
				>&nbsp;</td>
										  <td width="31"
				>&nbsp;</td>
										  <td width="31"
				>&nbsp;</td>
									  <!--										  <td width="37"
				> 
											<div align="center"> </div></td>-->
										  <td width="31"
				> 
											<div align="center"><%if isnumeric(m1_exibe) then 
									m1_exibe = m1_exibe*1	
									val_param = val_param*1	
									if m1_exibe<val_param then
										response.Write("<font color=#F00>"& m1_exibe&"</font>")	
									else
										response.Write(m1_exibe)									
									end if									
								else	
									response.Write(m1_exibe)
								end if	%> </div></td>
										  <td width="31"
				>&nbsp;</td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center">
                              <%if isnumeric(m2_exibe) then 
									m2_exibe = m2_exibe*1	
									val_param = val_param*1	
									if m2_exibe<val_param then
										response.Write("<font color=#F00>"& m2_exibe&"</font>")	
									else
										response.Write(m2_exibe)									
									end if									
								else	
									response.Write(m2_exibe)
								end if	%>
                            </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center">
                              <%if isnumeric(m3_exibe) then 
									m3_exibe = m3_exibe*1	
									val_param = val_param*1	
									if m3_exibe<val_param then
										response.Write("<font color=#F00>"& m3_exibe&"</font>")	
									else
										response.Write(m3_exibe)									
									end if									
								else	
									response.Write(m3_exibe)
								end if	%>
                            </div></td>
										  <td width="153"
				> <div align="center"> </div></td>
										  <td width="115"
				> <div align="center"> </div>
									  </td>
						</tr>
			<%
			end if
			end if

		RSprog.MOVENEXT
		wend
		%>
								<tr valign="bottom"> 
								  <td height="20" colspan="24" 
		> <div align="right"><font class="form_corpo"> 
        
        <% if etapa=6 or etapa=7 or etapa=8 or etapa=9 then
				Response.Write("T-Teste, PR-Prova, S - Simulado, MT=(T1+T2+T3+T4)/4, MP=(PR1+S), M3=((MTx1)+(MPx2))/3. <br>Para a Disciplina Portugu�s PR2 = Reda��o e M3=((MTx1)+(MPx2)+(PR2x2))/5.")			
		else
			Response.Write("AV-Avalia&ccedil;&otilde;es, MAV--M&eacute;dia das Avalia&ccedil;&otilde;es, SIM-Simulado, BAT-Bonus Atualidade  e M-M&eacute;dia")
		End if%>        
        
        </font></div></td>
								</tr>
					  </table>
<%
elseif notaFIL="TB_NOTA_M" then
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td width="125" class="tb_subtit"> <div align="left">Disciplina</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">AV1</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">AV2</div></td>
                          <td width="31" class="tb_subtit"> <div align="center">AV3</div></td>
<!--                          <td width="37" class="tb_subtit">&nbsp;</td>-->
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">AV4</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">AV5</div></td>
                          <td width="31" align="center" class="tb_subtit">SIM</td>
                          <td width="31" align="center" class="tb_subtit">MAV</td>
                          <td width="31" align="center" class="tb_subtit">BAT</td>
                          <td width="31" align="center" class="tb_subtit">BSI</td>
                          <!--                          <td width="37" class="tb_subtit">&nbsp;</td>-->
                          <td width="31" class="tb_subtit"> 
                            <div align="center">M1</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">Bon</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">M2</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">Rec</div></td>
                          <td width="31"  class="tb_subtit"> 
                            <div align="center">M3</div></td>
                          <td width="153" class="tb_subtit"> <div align="center">Alterado 
                              por</div></td>
                          <td width="115" class="tb_subtit"> <div align="center">Data/Hora</div></td>
                        </tr>
<!--                        <tr>
                          <td width="125" class="tb_subtit">&nbsp;</td> 
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">P</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center"> 
                              M</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">P</div></td>
                          <td width="37" class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37"  class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="36"  class="tb_subtit"> 
                            <div align="center">N</div></td>
                          <td width="37"  class="tb_subtit"> 
                            <div align="center">M</div></td>
                          <td width="176" class="tb_subtit">&nbsp;</td>
                          <td width="115" class="tb_subtit">&nbsp;</td>
                        </tr>
-->                        <%
		rec_lancado="sim"
		
				Set RSprog = Server.CreateObject("ADODB.Recordset")
				SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
				RSprog.Open SQLprog, CON0
		
		check=1
			
		while not RSprog.EOF
		
				materia=RSprog("CO_Materia")
				mae=RSprog("IN_MAE")
				fil=RSprog("IN_FIL")
				in_co=RSprog("IN_CO")
				nu_peso=RSprog("NU_Peso")
				ordem=RSprog("NU_Ordem_Boletim")
				
			peso_acumula=0
			m1_ac=0
			m2_ac=0			
			m3_ac=0
			m1_exibe=0
			m2_exibe=0
			m3_exibe=0
					
			if mae=TRUE THEN
			
			check=check+1
			
					Set RS1a = Server.CreateObject("ADODB.Recordset")
					SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia_Principal='"&materia&"'"
					RS1a.Open SQL1a, CON0
					
			if RS1a.EOF then
			
						Set RS1b = Server.CreateObject("ADODB.Recordset")
						SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
						RS1b.Open SQL1b, CON0
						
						no_materia=RS1b("NO_Materia")
						
						 if check mod 2 =0 then
						  cor = "tb_fundo_linha_par" 
						 else cor ="tb_fundo_linha_impar"
						  end if
							
								Set RSnFIL = Server.CreateObject("ADODB.Recordset")
								Set RS3 = Server.CreateObject("ADODB.Recordset")
								SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' and NU_Periodo="&periodo_check
								Set RS3 = CON_N.Execute(SQL_N)
						if RS3.EOF then
								va_av1="&nbsp;"
								va_av2="&nbsp;"
								va_av3="&nbsp;"								
								va_av4="&nbsp;"
								va_av5="&nbsp;"							
								va_mav="&nbsp;"
								va_sim="&nbsp;"
								va_bat="&nbsp;"
								va_bsi="&nbsp;"									
								va_m1="&nbsp;"
								va_bon="&nbsp;"
								va_m2="&nbsp;"
								va_rec="&nbsp;"
								va_m3="&nbsp;"
								data_grav="&nbsp;"
								hora_grav="&nbsp;"
								usuario_grav="&nbsp;"			
						else
						
								va_av1=RS3("VA_Av1")
								va_av2=RS3("VA_Av2")
								va_av3=RS3("VA_Av3")								
								va_av4=RS3("VA_Av4")
								va_av5=RS3("VA_Av5")						
								va_mav=RS3("VA_Mav")
								va_sim=RS3("VA_Sim")
								va_bat=RS3("VA_Bat")
								va_bsi=RS3("VA_Bsi")									
								va_m1=RS3("VA_Media1")
								va_bon=RS3("VA_Bonus")
								va_m2=RS3("VA_Media2")
								va_rec=RS3("VA_Rec")
								va_m3=RS3("VA_Media3")
								data_grav=RS3("DA_Ult_Acesso")
								hora_grav=RS3("HO_ult_Acesso")
								usuario_grav=RS3("CO_Usuario")
						end if

						
															
								
						if hora_grav="&nbsp;" or hora_grav="" or isnull(hora_grav) then
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
											
						if data_grav="&nbsp;" or data_grav="" or isnull(data_grav) then
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
						
						showapr="s"
						showprova="s"
						'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
						'data_inicio=""
						'va_faltas=""
						'		end if
						
						if usuario_grav="&nbsp;" then
						no_usuario=""
						else
								Set RS_pro = Server.CreateObject("ADODB.Recordset")
								SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
								RS_pro.Open SQL_pro, CON
							
							if RS_pro.eof then
								no_usuario= usuario_grav	
							else
								no_usuario=RS_pro("NO_Usuario")
							end if	
						end if
						%>
												<tr class="<%response.Write(cor)%>"> 
												  <td width="125"> 
													<%response.Write(no_materia)%>
												  </td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then							
													response.Write(va_av1)
													End IF							
													%>
													</div></td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_av2)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_av3)
													end if
													%>
													</div></td>
<!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													'response.Write(va_pt)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_av4)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_av5)
													end if
													%>
													</div></td>
												  <td width="31" align="center"
						><%if showapr="s" then					
													response.Write(va_sim)
													end if
													%></td>
												  <td width="31" align="center"
						><%if showprova="s" then					
													response.Write(va_mav)
													end if
													%></td>
												  <td width="31" align="center"
						><%if showprova="s" then					
													response.Write(va_bat)
													end if
													%><strong></strong></td>
												  <td width="31" align="center"
						><%if showprova="s" then					
													response.Write(va_bsi)
													end if
													%></td>
												  <!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													'response.Write(va_pp)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%

													if showprova="s" AND showapr="s" then					
														if isnumeric(va_m1) then
															if va_m1<val_param then
																response.Write("<font color=#F00>"& va_m1&"</font>")	
															else
																response.Write(va_m1)									
															end if									
														else	
															response.Write(va_m1)
														end if	
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_bon)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
								if isnumeric(va_m2) then
								    va_m2 = va_m2*1
									val_param = val_param*1
									if va_m2<val_param then
										response.Write("<font color=#F00>"& va_m2&"</font>")	
									else
										response.Write(va_m2)									
									end if									
								else	
									response.Write(va_m2)
								end if	
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_rec)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
								if isnumeric(va_m3) then
								    va_m3 = va_m3*1
									val_param = val_param*1
									if va_m3<val_param then
										response.Write("<font color=#F00>"& va_m3&"</font>")	
									else
										response.Write(va_m3)									
									end if									
								else	
									response.Write(va_m3)
								end if	
													end if
													%>
													</div></td>
												  <td width="153"
						> <div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then
													response.Write(no_usuario)
													end if
													%>
													</div></td>
												  <td width="115"
						> <div align="center"> 
													<%
													if showprova="s" AND showapr="s" then							
													response.Write(data_inicio)
													End if
													%></div>
												  </td>
												</tr>
					<%
			else

			
			
				 if check mod 2 =0 then
				  cor = "tb_fundo_linha_par" 
				 else cor ="tb_fundo_linha_impar"
				  end if
			
						Set RS1b = Server.CreateObject("ADODB.Recordset")
						SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
						RS1b.Open SQL1b, CON0
						
				no_materia=RS1b("NO_Materia")
					
						Set RSnFIL = Server.CreateObject("ADODB.Recordset")
						Set RS3 = Server.CreateObject("ADODB.Recordset")
						SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' and NU_Periodo="&periodo_check						
						Set RS3 = CON_N.Execute(SQL_N)
						if RS3.EOF then
								va_av1="&nbsp;"
								va_av2="&nbsp;"
								va_av3="&nbsp;"								
								va_av4="&nbsp;"
								va_av5="&nbsp;"							
								va_mav="&nbsp;"
								va_sim="&nbsp;"
								va_bat="&nbsp;"
								va_bsi="&nbsp;"								
								va_m1="&nbsp;"
								va_bon="&nbsp;"
								va_m2="&nbsp;"
								va_rec="&nbsp;"
								va_m3="&nbsp;"
								data_grav="&nbsp;"
								hora_grav="&nbsp;"
								usuario_grav="&nbsp;"			
						else
						
								va_av1=RS3("VA_Av1")
								va_av2=RS3("VA_Av2")
								va_av3=RS3("VA_Av3")								
								va_av4=RS3("VA_Av4")
								va_av5=RS3("VA_Av5")						
								va_mav=RS3("VA_Mav")
								va_sim=RS3("VA_Sim")
								va_bat=RS3("VA_Bat")
								va_bsi=RS3("VA_Bsi")								
								va_m1=RS3("VA_Media1")
								va_bon=RS3("VA_Bonus")
								va_m2=RS3("VA_Media2")
								va_rec=RS3("VA_Rec")
								va_m3=RS3("VA_Media3")
								data_grav=RS3("DA_Ult_Acesso")
								hora_grav=RS3("HO_ult_Acesso")
								usuario_grav=RS3("CO_Usuario")
						end if
						
				if hora_grav="&nbsp;" or hora_grav="" or isnull(hora_grav) then
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
									
				if data_grav="&nbsp;" or data_grav="" or isnull(data_grav)  then
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
				
				showapr="s"
				showprova="s"
				'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
				'data_inicio=""
				'va_faltas=""
				'		end if
				
				if usuario_grav="&nbsp;" then
				no_usuario=""
				else
						Set RS_pro = Server.CreateObject("ADODB.Recordset")
						SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
						RS_pro.Open SQL_pro, CON
				
							if RS_pro.eof then
								no_usuario= usuario_grav	
							else
								no_usuario=RS_pro("NO_Usuario")
							end if	
				end if
				%>
										<tr class="<%response.Write(cor)%>"> 
										  <td width="125"> 
											<%response.Write(no_materia)%>
										  </td>
										  <td width="31"> 
											<div align="center"> 
											  <%
											if showapr="s" then							
											response.Write(va_av1)
											End IF							
											%>
											</div></td>
										  <td width="31"> 
											<div align="center"> 
											  <%
											if showapr="s" then					
											response.Write(va_av2)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showapr="s" then					
											response.Write(va_av3)
											end if
											%>
											</div></td>
<!--										  <td width="37"
				> 
											<div align="center"> 
											  <%
											if showapr="s" then					
											'response.Write(va_pt)
											end if
											%>
											</div></td>-->
										  <td width="31"
				> 

											<div align="center"> 
											  <%
											if showapr="s" then					
											response.Write(va_av4)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showapr="s" then					
											response.Write(va_av5)
											end if
											%>
											</div></td>
										  <td align="center"
						><%if showapr="s" then					
													response.Write(va_sim)
													end if
													%></td>
										  <td align="center"
						><%if showprova="s" then					
													response.Write(va_mav)
													end if
													%></td>
										  <td align="center"
						><%if showprova="s" then					
													response.Write(va_bat)
													end if
													%>
										    <strong></strong></td>
										  <td align="center"
						><%if showprova="s" then					
													response.Write(va_bsi)
													end if
													%></td>
										  <!--										  <td width="37"
				> 
											<div align="center"> 
											  <%
											if showprova="s" then					
											'response.Write(va_pp)
											end if
											%>
											</div></td>-->
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_m1)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_bon)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_m2)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%

											if showprova="s" AND showapr="s" then					
											response.Write(va_rec)
											end if
											%>
											</div></td>
										  <td width="31"
				> 
											<div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_m3)
											end if
											%>
											</div></td>
										  <td width="153"
				> <div align="center"> 
											  <%
											if showprova="s" AND showapr="s" then
											response.Write(no_usuario)
											end if
											%>
											</div></td>
										  <td width="115"
				> <div align="center"> 
											<%
											if showprova="s" AND showapr="s" then							
											response.Write(data_inicio)
											End if
											%></div>
										  </td>
										</tr>
			<%
			peso_acumula=0
			acumula_m1=0
			m2_ac=0			
			m3_ac=0
			m1_exibe=0
			m2_exibe=0
			m3_exibe=0
			
			
				while not RS1a.EOF
				
						materia_fil=RS1a("CO_Materia")
					
								Set RS1b = Server.CreateObject("ADODB.Recordset")
								SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia_fil&"'"
								RS1b.Open SQL1b, CON0
								
						no_materia_fil=RS1b("NO_Materia")
						
						Set RSpa = Server.CreateObject("ADODB.Recordset")
						SQLpa= "SELECT * FROM TB_Programa_Aula where CO_Materia='"&materia_fil&"' AND CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"'"
						RSpa.Open SQLpa, CON0
												
						nu_peso_fil=RSpa("NU_Peso")		
						if isnull(nu_peso_fil) or nu_peso_fil="" then
							nu_peso_fil=1
						end if												
							
								Set RSnFIL = Server.CreateObject("ADODB.Recordset")
								Set RS3 = Server.CreateObject("ADODB.Recordset")
								SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia_Principal = '"& materia &"' AND CO_Materia = '"& materia_fil &"' and NU_Periodo="&periodo_check
								Set RS3 = CON_N.Execute(SQL_N)
						if RS3.EOF then
								va_av1="&nbsp;"
								va_av2="&nbsp;"
								va_av3="&nbsp;"								
								va_av4="&nbsp;"
								va_av5="&nbsp;"							
								va_mav="&nbsp;"
								va_sim="&nbsp;"
								va_bat="&nbsp;"
								va_bsi="&nbsp;"								
								va_m1="&nbsp;"
								va_bon="&nbsp;"
								va_m2="&nbsp;"
								va_rec="&nbsp;"
								va_m3="&nbsp;"
								data_grav="&nbsp;"
								hora_grav="&nbsp;"
								usuario_grav="&nbsp;"			
						else
						
								va_av1=RS3("VA_Av1")
								va_av2=RS3("VA_Av2")
								va_av3=RS3("VA_Av3")								
								va_av4=RS3("VA_Av4")
								va_av5=RS3("VA_Av5")						
								va_mav=RS3("VA_Mav")
								va_sim=RS3("VA_Sim")
								va_bat=RS3("VA_Bat")
								va_bsi=RS3("VA_Bsi")
								va_m1=RS3("VA_Media1")
								va_bon=RS3("VA_Bonus")
								va_m2=RS3("VA_Media2")
								va_rec=RS3("VA_Rec")
								va_m3=RS3("VA_Media3")
								data_grav=RS3("DA_Ult_Acesso")
								hora_grav=RS3("HO_ult_Acesso")
								usuario_grav=RS3("CO_Usuario")
						end if
						
						if isnull(va_m1) or va_m1="" or va_m1="&nbsp;" then
							'va_m1=0
							sem_media1="s"
						else
							sem_media1="n"
							m1_ac=m1_ac+(va_m1*nu_peso_fil)															
						end if

						if isnull(va_m2) or va_m2="" or va_m2="&nbsp;" then
							'va_m2=0
							sem_media2="s"
						else
							sem_media2="n"	
							m2_ac=m2_ac+(va_m2*nu_peso_fil)							
						end if
						
						if isnull(va_m3) or va_m3="" or va_m3="&nbsp;" then
							'va_m3=0
							sem_media3="s"
						else
							sem_media3="n"		
							m3_ac=m3_ac+(va_m3*nu_peso_fil)					
						end if												
						
							'peso_acumula=peso_acumula+nu_peso_fil
						peso_acumula=nu_peso_fil								
													
								
						if hora_grav="&nbsp;" or hora_grav="" or isnull(hora_grav)  then
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
											
						if data_grav="&nbsp;" or data_grav="" or isnull(data_grav) then
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
						
						showapr="s"
						showprova="s"
						'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
						'data_inicio=""
						'va_faltas=""
						'		end if
						
						if usuario_grav="&nbsp;" then
						no_usuario=""
						else
								Set RS_pro = Server.CreateObject("ADODB.Recordset")
								SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
								RS_pro.Open SQL_pro, CON
						
							if RS_pro.eof then
								no_usuario= usuario_grav	
							else
								no_usuario=RS_pro("NO_Usuario")
							end if	
						end if
						%>
												
											<tr class="<%response.Write(cor)%>"> 
											  <td width="125">&nbsp;&nbsp;&nbsp;
												  <%response.Write(no_materia_fil)%>
											  </td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then							
													response.Write(va_av1)
													End IF							
													%>
													</div></td>
												  <td width="31"> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_av2)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_av3)
													end if
													%>
													</div></td>
<!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													'response.Write(va_pt)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_av4)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showapr="s" then					
													response.Write(va_av5)
													end if
													%>
													</div></td>
												  <td align="center"
						><%if showapr="s" then					
													response.Write(va_sim)
													end if
													%></td>
												  <td align="center"
						><%if showprova="s" then					
													response.Write(va_mav)
													end if
													%></td>
												  <td align="center"
						><%if showprova="s" then					
													response.Write(va_bat)
													end if
													%></td>
												  <td align="center"
						><%if showprova="s" then					
													response.Write(va_bsi)
													end if
													%></td>
											  <!--												  <td width="37"
						> 
													<div align="center"> 
													  <%
													if showprova="s" then					
													'response.Write(va_pp)
													end if
													%>
													</div></td>-->
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m1)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_bon)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m2)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_rec)
													end if
													%>
													</div></td>
												  <td width="31"
						> 
													<div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m3)
													end if
													%>
													</div></td>
												  <td width="153"
						> <div align="center"> 
													  <%
													if showprova="s" AND showapr="s" then
													response.Write(no_usuario)
													end if
													%>
													</div></td>
												  <td width="115"
						> <div align="center"> 
													<%
													if showprova="s" AND showapr="s" then							
													response.Write(data_inicio)
													End if
													%></div>
											  </td>
						</tr>
				<%
				RS1a.movenext
				wend
						if	sem_media1="s" then
							m1_exibe="&nbsp;"
						else
							m1_exibe=m1_ac/peso_acumula
							decimo = m1_exibe - Int(m1_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m1_exibe) + 1
									m1_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m1_exibe)
									m1_exibe=nota_arredondada					
								End If
							if m1_exibe>100 then
								m1_exibe=100
							end if
							m1_exibe= formatNumber(m1_exibe,0)							
						end if	
							
						if	sem_media2="s" then	
							m2_exibe="&nbsp;"
						else												
							m2_exibe=m2_ac/peso_acumula
							decimo = m2_exibe - Int(m2_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m2_exibe) + 1
									m2_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m2_exibe)
									m2_exibe=nota_arredondada					
								End If
							if m2_exibe>100 then
								m2_exibe=100
							end if								
							m2_exibe= formatNumber(m2_exibe,0)							
						end if	
						
						if	sem_media3="s" then
							m3_exibe="&nbsp;"
						else							
							m3_exibe=m3_ac/peso_acumula
								decimo = m3_exibe - Int(m3_exibe)
									If decimo >= 0.5 Then
										nota_arredondada = Int(m3_exibe) + 1
										m3_exibe=nota_arredondada
									Else
										nota_arredondada = Int(m3_exibe)
										m3_exibe=nota_arredondada					
									End If
								if m3_exibe>100 then
									m3_exibe=100
								end if										
								m3_exibe= formatNumber(m3_exibe,0)									
						end if														
				
				%>
									<tr class="tb_fundo_linha_media"> 
									  <td width="125">&nbsp;&nbsp;&nbsp; M&eacute;dia</td>
										  <td width="31"> 
											<div align="center"></div></td>
										  <td width="31"> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
<!--										  <td width="37"
				> 
											<div align="center"> </div></td>-->
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				>&nbsp;</td>
										  <td width="31"
				>&nbsp;</td>
										  <td width="31"
				>&nbsp;</td>
										  <td width="31"
				>&nbsp;</td>
									  <!--										  <td width="37"
				> 
											<div align="center"> </div></td>-->
										  <td width="31"
				> 
											<div align="center"><%if isnumeric(m1_exibe) then 
									m1_exibe = m1_exibe*1	
									val_param = val_param*1	
									if m1_exibe<val_param then
										response.Write("<font color=#F00>"& m1_exibe&"</font>")	
									else
										response.Write(m1_exibe)									
									end if									
								else	
									response.Write(m1_exibe)
								end if	%> </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center">
                              <%if isnumeric(m2_exibe) then 
									m2_exibe = m2_exibe*1	
									val_param = val_param*1	
									if m2_exibe<val_param then
										response.Write("<font color=#F00>"& m2_exibe&"</font>")	
									else
										response.Write(m2_exibe)									
									end if									
								else	
									response.Write(m2_exibe)
								end if	%>
                            </div></td>
										  <td width="31"
				> 
											<div align="center"> </div></td>
										  <td width="31"
				> 
											<div align="center">
                              <%if isnumeric(m3_exibe) then 
									m3_exibe = m3_exibe*1	
									val_param = val_param*1	
									if m3_exibe<val_param then
										response.Write("<font color=#F00>"& m3_exibe&"</font>")	
									else
										response.Write(m3_exibe)									
									end if									
								else	
									response.Write(m3_exibe)
								end if	%>
                            </div></td>
										  <td width="153"
				> <div align="center"> </div></td>
										  <td width="115"
				> <div align="center"> </div>
									  </td>
						</tr>
			<%
			end if
			end if

		RSprog.MOVENEXT
		wend
		%>
								<tr valign="bottom"> 
								  <td height="20" colspan="23" 
		> <div align="right"><font class="form_corpo"> 
        
        <% if etapa=6 or etapa=7 or etapa=8 or etapa=9 then
				Response.Write("T-Teste, PR-Prova, S - Simulado, MT=(T1+T2+T3+T4)/4, MP=(PR1+S), M3=((MTx1)+(MPx2))/3. <br>Para a Disciplina Portugu�s PR2 = Reda��o e M3=((MTx1)+(MPx2)+(PR2x2))/5.")			
		else
			Response.Write("AV-Avalia&ccedil;&otilde;es, MAV-M&eacute;dia das Avalia&ccedil;&otilde;es, SIM-Simulado, BAT-Bonus Atualidade  e M-M&eacute;dia")
		End if%>        
        
        </font></div></td>
								</tr>
					  </table>                                            
<%end if
end if%>
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