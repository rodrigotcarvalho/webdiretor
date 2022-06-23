<!--#include file="../../inc/caminhos.asp"-->
<!--#include file="../../inc/funcoes.asp"-->
<!--#include file="../../inc/funcoes2.asp"-->


<%
nivel=2
tp=session("tp")
nome = session("nome") 
co_user = session("co_user")
opt=request.QueryString("opt")

if opt="1" then
periodo_check=request.form("periodo")
cod= Session("aluno_selecionado")
else
cod= Session("aluno_selecionado")
periodo_check=1
end if
cod= Session("aluno_selecionado")

obr=cod&"?"&periodo_check

 	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0

 	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1
	
	Set CON2 = Server.CreateObject("ADODB.Connection") 
	ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON2.Open ABRIR2	



	SQL2 = "select * from TB_Usuario where CO_Usuario = " & cod 
	set RS2 = CON.Execute (SQL2)
	
nome_aluno= RS2("NO_Usuario")

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



%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Web Família</title>
<link href="../../estilo.css" rel="stylesheet" type="text/css" />
<script type="text/JavaScript">
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
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

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

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
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
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
function submitfuncao()  
{
   var f=document.forms[0]; 
      f.submit(); 
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>

<body onLoad="MM_preloadImages(<%response.Write(swapload)%>)">
<form action="index.asp?opt=1" method="post"><table width="1000" height="1038" border="0" align="center" cellpadding="0" cellspacing="0" class="tabela_principal">
  <%
			ano = DatePart("yyyy", now)
			mes = DatePart("m", now) 
			dia = DatePart("d", now) 
			hora = DatePart("h", now) 
			min = DatePart("n", now) 
select case mes
 case 1 
 mes = "janeiro"
 case 2 
 mes = "fevereiro"
 case 3 
 mes = "março"
 case 4
 mes = "abril"
 case 5
 mes = "maio"
 case 6 
 mes = "junho"
 case 7
 mes = "julho"
 case 8 
 mes = "agosto"
 case 9 
 mes = "setembro"
 case 10 
 mes = "outubro"
 case 11 
 mes = "novembro"
 case 12 
 mes = "dezembro"
end select

data = dia &" / "& mes &" / "& ano
data= FormatDateTime(data,1) 			

			horario = hora & ":"& min%>
  <tr>
    <td height="998"><table width="200" height="998" border="0" cellpadding="0" cellspacing="0">
          <!--DWLayoutTable-->
          <tr> 
            <td height="130" colspan="3">
              <%call cabecalho(nivel)%>
            </td>
          </tr>
          <tr class="tabela_menu"> 
            <td width="172" height="144" rowspan="4" valign="top" class="tabela_menu"><p><img src="../../img/baner_resto.jpg" width="171" height="38" /></p>
              <% call menu_lateral(nivel)%>
              <p>&nbsp;</p></td>
            <td width="640" height="12" nowrap="nowrap"><p class="style1">&nbsp;&nbsp;Ol&aacute; 
                <span class="style2">
                <%response.Write(nome)%>
                </span> , &uacute;ltimo acesso dia 
                <% Response.Write(session("dia_t")) %>
                &agrave;s 
                <% Response.Write(session("hora_t")) %>
              </p></td>
            <td width="188"><p align="right" class="style1"> 
                <%response.Write(data)%>
              </p></td>
          </tr>
          <tr class="tabela_menu"> 
            <td height="5" colspan="2"><p><img src="../../img/linha-pontilhada_grande.gif" alt="" width="828" height="5" /></p></td>
          </tr>
      <tr class="tabela_menu">
        <td height="19" colspan="2">&nbsp;</td>
      </tr>		  
          <tr class="tabela_menu"> 
            <td height="832" colspan="2" valign="top"><img src="../../img/avaliacoes.jpg" width="700" height="30"> 
              <table width="100%" border="0" align="center" cellspacing="0" class="tb_corpo"
>
                <tr> 
                  <td width="684" class="tb_tit"
>Dados Escolares</td>
                  <td width="140" class="tb_tit"
> </td>
                </tr>
                <tr> 
                  <td height="10"> <table width="100%" border="0" cellspacing="0">
                      <tr> 
                        <td width="19%" height="10"> <div align="right"><font class="style3"> 
                            Matr&iacute;cula: </font></div></td>
                        <td width="9%" height="10"><font class="style1"> 
                          <input name="cod" type="hidden" value="<%=cod%>">
                          <%response.Write(cod)%>
                          </font></td>
                        <td width="6%" height="10"> <div align="right"><font class="style3"> 
                            Nome: </font></div></td>
                        <td width="66%" height="10"><font class="style1"> 
                          <%response.Write(nome_aluno)%>
                          <input name="nome2" type="hidden" class="textInput" id="nome2"  value="<%response.Write(nome_aluno)%>" size="75" maxlength="50">
                          &nbsp;</font></td>
                      </tr>
                    </table></td>
                  <td rowspan="2" valign="top"><div align="center"> </div></td>
                </tr>
                <tr> 
                  <td height="10" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="2"><table width="100%" border="0" cellspacing="0">
                      <tr class="style3"> 
                        <td width="34" height="10"> <div align="center"> 
                            <%					  
call GeraNomes("PORT",unidade,curso,etapa,CON0)
no_unidade = session("no_unidade")
no_curso = session("no_curso")
no_etapa= session("no_etapa")
%>
                            ANO</div></td>
                        <td width="74" height="10"> <div align="center">MATR&Iacute;CULA</div></td>
                        <td width="96" height="10"> <div align="center">CANCELAMENTO</div></td>
                        <td width="83" height="10"> <div align="center"> SITUA&Ccedil;&Atilde;O</div></td>
                        <td width="81" height="10"> <div align="center">UNIDADE</div></td>
                        <td width="91" height="10"> <div align="center">CURSO</div></td>
                        <td width="63" height="10"> <div align="center"> ETAPA</div></td>
                        <td width="66" height="10"> <div align="center">TURMA</div></td>
                        <td width="81" height="10"> <div align="center">CHAMADA</div></td>
                        <td width="137"> <div align="center">PER&Iacute;ODO</div></td>
                      </tr>
                      <tr class="tb_corpo"
> 
                        <td width="34" height="10"> <div align="center"> <font class="style1"> 
                            <%response.Write(ano_aluno)%>
                            </font></div></td>
                        <td width="74" height="10"> <div align="center"> <font class="style1"> 
                            <%response.Write(rematricula)%>
                            </font></div></td>
                        <td width="96" height="10"> <div align="center"> <font class="style1"> 
                            <%response.Write(encerramento)%>
                            </font></div></td>
                        <td width="83" height="10"> <div align="center"> <font class="style1"> 
                            <%
					
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
		RSCONTST.Open SQLCONTST, CON0
							
				no_situacao = RSCONTST("TX_Descricao_Situacao")	
					response.Write(no_situacao)%>
                            </font></div></td>
                        <td width="81" height="10"> <div align="center"> <font class="style1"> 
                            <%response.Write(no_unidade)%>
                            </font></div></td>
                        <td width="91" height="10"> <div align="center"> <font class="style1"> 
                            <%response.Write(no_curso)%>
                            </font></div></td>
                        <td width="63" height="10"> <div align="center"> <font class="style1"> 
                            <%response.Write(no_etapa)%>
                            </font></div></td>
                        <td width="66" height="10"> <div align="center"> <font class="style1"> 
                            <%response.Write(turma)%>
                            </font></div></td>
                        <td width="81" height="10"> <div align="center"> <font class="style1"> 
                            <%response.Write(cham)%>
                            </font></div></td>
                        <td width="137"> <div align="center"> 
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
>Avalia&ccedil;&otilde;es Progressivas</td>
                </tr>
                <tr> 
                  <td colspan="2"> 
                    <%		Set RS_tb = Server.CreateObject("ADODB.Recordset")
		SQL_tb = "SELECT * FROM TB_Da_Aula WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso&"' AND CO_Etapa ='"& etapa&"' AND CO_Turma ='"& turma&"'"
		RS_tb.Open SQL_tb, CON2

if RS_tb.eof then
%><div align="center"> <font class="style1"> 
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
                        <td width="<%response.Write(width)%>" rowspan="2" class="style3"
> <div align="left"><strong>Disciplina</strong></div></td>
                        <td width="30" rowspan="2" class="style3"
> <div align="center">F</div></td>
                        <td colspan="2" class="style3"
><div align="center">APR1</div></td>
                        <td colspan="2" class="style3"
><div align="center">APR2</div></td>
                        <td colspan="2" class="style3"
><div align="center">APR3</div></td>
                        <td colspan="2" class="style3"
><div align="center">APR4</div></td>
                        <td colspan="2" class="style3"
><div align="center">APR5</div></td>
                        <td colspan="2" class="style3"
><div align="center">APR6</div></td>
                        <td colspan="2" class="style3"
><div align="center">TEC1</div></td>
                        <td colspan="2" class="style3"
><div align="center">TEC2</div></td>
                        <td width="30" rowspan="2" class="style3"
> <div align="center">SAPR</div></td>
                        <td width="30" rowspan="2" class="style3"
> <div align="center">PR</div></td>
                        <td width="30" rowspan="2" class="style3"
> <div align="center">MP</div></td>
<% if periodo_check=2 then%>
                          <td width="30" rowspan="2" class="style3"
> <div align="center">EC1</div></td>
<%end if%>
 <!--                         <td width="115" rowspan="2" class="style3"
> <div align="center">Alterado em</div></td>-->
                      </tr>
                      <tr> 
                        <td width="30" class="style3"
> <div align="center">N</div></td>
                        <td width="30" class="style3"
> <div align="center">P</div></td>
                        <td width="30" class="style3"
> <div align="center">N</div></td>
                        <td width="30" class="style3"
> <div align="center">P</div></td>
                        <td width="30" class="style3"
> <div align="center">N</div></td>
                        <td width="30" class="style3"
> <div align="center">P</div></td>
                        <td width="30" class="style3"
> <div align="center">N</div></td>
                        <td width="30" class="style3"
> <div align="center">P</div></td>
                        <td width="30" class="style3"
> <div align="center">N</div></td>
                        <td width="30" class="style3"
> <div align="center">P</div></td>
                        <td width="30" class="style3"
> <div align="center">N</div></td>
                        <td width="30" class="style3"
> <div align="center">P</div></td>
                        <td width="30" class="style3"
> <div align="center">N</div></td>
                        <td width="30" class="style3"
> <div align="center">P</div></td>
                        <td width="30" class="style3"
> <div align="center">N</div></td>
                        <td width="30" class="style3"
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

		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL_WF = "SELECT * FROM TB_Autoriza_WF WHERE NU_Unidade="&unidade&" AND CO_Curso='"&curso&"' and CO_Etapa='"&etapa&"'"
		RS4.Open SQL_WF, CON	

	
co_apr1=RS4("CO_apr1")
co_apr2=RS4("CO_apr2")
co_apr3=RS4("CO_apr3")
co_apr4=RS4("CO_apr4")
co_prova1=RS4("CO_prova1")
co_prova2=RS4("CO_prova2")
co_prova3=RS4("CO_prova3")
co_prova4=RS4("CO_prova4")	
		
if periodo_check=1 then		
		if co_apr1="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova1="D"then
		showprova="n"
		else 
		showprova="s"
		end if
elseif periodo_check=2 then		
		if co_apr2="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova2="D"then
		showprova="n"
		else 
		showprova="s"
		end if					
elseif periodo_check=3 then		
		if co_apr3="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova3="D"then
		showprova="n"
		else 
		showprova="s"
		end if
elseif periodo_check=4 then		
		if co_apr4="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova4="D"then
		showprova="n"
		else 
		showprova="s"
		end if
end if											
		
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

if (isnull(va_apr1) OR va_apr1="") and (ISNULL(va_apr2) OR va_apr2="")and (ISNULL(va_apr3) OR va_apr3="")and (ISNULL(va_apr4) OR va_apr4="")and (ISNULL(va_apr5)  OR va_apr5="")and (ISNULL(va_apr6) OR  va_apr6="") and (ISNULL(va_apr7) OR va_apr7="")and (ISNULL(va_apr8) OR va_apr8="")and (ISNULL(va_sapr) OR va_sapr="")  then
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
<% if periodo_check=2 then%>
                          <td width="30"> 
						  <div align="center"><font class="style1"> 
							<%
							if showprova="s" then
							response.Write(va_bon)							
							else					
							end if%></font></div>
							</td>
<%end if%>
<!-- 							
                        <td width="115"
> <div align="center"><font class="style1"> 
                            <%
							'	if showprova="n" AND showapr="n" then
							'else							
							'response.Write(data_inicio)
							'End if
							%>
                            </font></div></td> -->
                      </tr>
                      <%check=check+1
RSprog.MOVENEXT
wend%>
                       <tr valign="bottom"> 
 <!--                       <td height="20" colspan="22" > -->
 <%if periodo_check=2 then%>							  
                          <td height="20" colspan="22" 
> <div align="right"><font class="style1"> F-Faltas , N-Nota Apr, P-Peso Apr SAPR–Média 
                              das Aprs, PR-Prova, MP–Média Período e ECE1-1ª Etapa Complementar de Estudos</font></div></td>							  
<%else%>							
                          <td height="20" colspan="21" 
> <div align="right"><font class="style1"> F-Faltas , N-Nota Apr, P-Peso Apr, SAPR–Média 
                              das Aprs, PR-Prova e MP–Média Período</font></div></td>
<%end if%>
                      </tr>
                      <tr> 
                        <td colspan="22" 
>&nbsp;</td>
                      </tr>
                      <tr> 
                        <td colspan="22" 
><table width="200" height="20" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td class="tb_tit"><div align="center"><a href="#" class="impressao" onClick="MM_openBrWindow('imprime.asp?obr=<%=obr%>','','menubar=yes,width=1000,height=450')">Versão 
                                  para impressão</a></div></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table>
<%end if%>					
					</td>
                </tr>
              </table></td>
          </tr>
        </table></td>
  </tr>
  <tr>
    <td width="1000" height="40"><img src="../../img/rodape.jpg" width="1000" height="40" /></td>
  </tr>
</table>
</form>
</body>
</html>
