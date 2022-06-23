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
'periodo_check=request.form("periodo")
cod= Session("aluno_selecionado")
else
cod= Session("aluno_selecionado")
'periodo_check=1
end if
cod= Session("aluno_selecionado")

obr=cod

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
    <td height="990" valign="top"><table width="200" height="990" border="0" cellpadding="0" cellspacing="0">
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
            <td height="832" colspan="2" valign="top"> <div align="left"><img src="../../img/boletim.jpg" width="700" height="30"> 
              </div>
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
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td bgcolor="#FFFFFF">&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td height="10" colspan="2" class="tb_tit"
>Boletim</td>
                </tr>
                <tr> 
                  <td colspan="2"> 
                    <%		Set RS_tb = Server.CreateObject("ADODB.Recordset")
		SQL_tb = "SELECT * FROM TB_Da_Aula WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso&"' AND CO_Etapa ='"& etapa&"' AND CO_Turma ='"& turma&"'"
		RS_tb.Open SQL_tb, CON2

if RS_tb.eof then
%><div align="center"> <font class="style1"> 
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
                        <td width="111" rowspan="2" class="style3"
> <div align="left"><strong>Disciplina</strong></div></td>
                        <td colspan="5" class="style3"
> <div align="center">TRIMESTRE 1</div></td>
                        <td colspan="5" class="style3"
><div align="center">TRIMESTRE 2</div></td>
                        <td colspan="5" class="style3"
><div align="center">TRIMESTRE 3</div></td>
                        <td width="30" rowspan="2" class="style3"
> <div align="center">RA </div>
                          <div align="center"></div></td>
                        <td colspan="4" class="style3"
><div align="center">ETAPA COMPLEMENTAR</div></td>
                        <td width="30" rowspan="2" class="style3"
> <div align="center">RF</div></td>
                      </tr>
                      <tr> 
                        <td width="30" class="style3"
> <div align="center">SAPR</div></td>
                        <td width="30" class="style3"
> <div align="center">PR</div></td>
                        <td width="30" class="style3"
> <div align="center">MP</div></td>
                        <td width="30" class="style3"
> <div align="center">MC</div></td>
                        <td width="30" class="style3"
> <div align="center">F</div></td>
                        <td width="30" class="style3"
> <div align="center">SAPR</div></td>
                        <td width="30" class="style3"
> <div align="center">PR</div></td>
                        <td width="30" class="style3"
> <div align="center">MP</div></td>
                        <td width="30" class="style3"
> <div align="center">MC*</div></td>
                        <td width="30" class="style3"
> <div align="center">F</div></td>
                        <td width="30" class="style3"
> <div align="center">SAPR</div></td>
                        <td width="30" class="style3"
> <div align="center">PR</div></td>
                        <td width="30" class="style3"
> <div align="center">MP</div></td>
                        <td width="30" class="style3"
> <div align="center">MC</div></td>
                        <td width="30" class="style3"
> <div align="center">F</div></td>
                        <td width="38" class="style3"
> <div align="center">SAPR</div></td>
                        <td width="38" class="style3"
> <div align="center">PR</div></td>
                        <td width="38" class="style3"
> <div align="center">MP</div></td>
                        <td width="38" class="style3"
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
va_sapr1=""
va_pr1=""
va_te1=""
pr1=""
va_me1=""
va_mc1=""
va_faltas1=""
va_sapr2=""
va_pr2=""
va_te2=""
pr2=""
va_me2=""
va_mc2=""
va_faltas2=""
va_sapr3=""
va_pr3=""
va_te3=""
pr3=""
va_me3=""
va_mc3=""
va_faltas3=""
va_sapr4=""
va_pr4=""
va_me4=""
va_mc4=""
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
		
'if periodo_check=1 then		
		if co_apr1="D"then
		showapr1="n"
		else 
		showapr1="s"
		end if
		if co_prova1="D"then
		showprova1="n"
		else 
		showprova1="s"
		end if
'elseif periodo_check=2 then	
		if co_apr2="D"then
		showapr2="n"
		else 
		showapr2="s"
		end if
		if co_prova2="D"then
		showprova2="n"
		else 
		showprova2="s"
		end if					
'elseif periodo_check=3 then		
		if co_apr3="D"then
		showapr3="n"
		else 
		showapr3="s"
		end if
		if co_prova3="D"then
		showprova3="n"
		else 
		showprova3="s"
		end if
'elseif periodo_check=4 then		
		if co_apr4="D"then
		showapr4="n"
		else 
		showapr4="s"
		end if
		if co_prova4="D"then
		showprova4="n"
		else 
		showprova4="s"
		end if
'end if											
		
		
				
end if

if va_me1="" or isnull(va_me1)then
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
	
if va_mc1="" or isnull(va_mc1)then
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
if va_me2="" or isnull(va_me2)then
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
	
if va_mc2="" or isnull(va_mc2)then
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
if va_me3="" or isnull(va_me3)then
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
	
if va_mc3="" or isnull(va_mc3)then
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
if va_me4="" or isnull(va_me4)then
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
	
if va_mc4="" or isnull(va_mc4)then
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
                        <td width="111"
><font class="style1"> 
                          <%response.Write(no_materia)%>
                          </font></td>
                        <td width="30"
> <div align="center"><font class="style1"> 
                            <%
							if showapr1="s" and showprova1="s" then
	if va_sapr1="" or isnull(va_sapr1) then
	else						
	va_sapr1 = formatNumber(va_sapr1,1)
	end if												
							response.Write(va_sapr1)
							end if
							%>
                            </font></div></td>
                        <td width="30"
> <div align="center"><font class="style1"> 
                            <%
							if showapr1="s" and showprova1="s" then
	if pr1="" or isnull(pr1) then
	else							
	pr1 = formatNumber(pr1,1)												
							response.Write(pr1)
	end if							
							end if
							%>
                            </font></div></td>
                        <td width="30"
> <div align="center"><font class="style1"> 
                            <%
							if showapr1="s" and showprova1="s" then					
							response.Write(va_me1)
							end if
							%>
                            </font></div></td>
                        <td width="30"
> <div align="center"><font class="style1"> 
                            <%
							if showapr1="s" and showprova1="s" then					
							response.Write(va_mc1)
							end if
							%>
                            </font></div></td>
                        <td width="30"
> <div align="center"><font class="style1"> 
                            <%
							if showapr1="s" and showprova1="s" then
							response.Write(va_faltas1)
							else
							end if
							%>
                            </font></div></td>
                        <td width="30"
> <div align="center"><font class="style1"> 
                            <%
							if showapr2="s" and showprova2="s" then
	if va_sapr2="" or isnull(va_sapr2) then
	else							
	va_sapr2 = formatNumber(va_sapr2,1)
													
							response.Write(va_sapr2)
	end if							
							end if
							%>
                            </font></div></td>
                        <td width="30"
> <div align="center"><font class="style1"> 
                            <%
							if showapr2="s" and showprova2="s" then
	if pr2="" or isnull(pr2) then
	else							
	pr2 = formatNumber(pr2,1)												
							response.Write("&nbsp;"&pr2)
	end if
							end if
							%>
                            </font></div></td>
                        <td width="30"
> <div align="center"><font class="style1"> 
                            <%
							if showapr2="s" and showprova2="s" then					
							response.Write(va_me2)
							end if
							%>
                            </font></div></td>
                        <td width="30"
> <div align="center"><font class="style1"> 
                            <%
							if showapr2="s" and showprova2="s" then					
							response.Write(va_mc2)
							end if
							%>
                            </font></div></td>
                        <td width="30"
> <div align="center"><font class="style1"> 
                            <%
							if showapr2="s" and showprova2="s" then
							response.Write(va_faltas2)
							else
							end if
							%>
                            </font></div></td>
                        <td width="30"
> <div align="center"><font class="style1"> 
                            <%
							if showapr3="s" and showprova3="s" then
	if va_sapr3="" or isnull(va_sapr3) then
	else							
	va_sapr3 = formatNumber(va_sapr3,1)
													
							response.Write(va_sapr3)
	end if	
							end if
							%>
                            </font></div></td>
                        <td width="30"
> <div align="center"><font class="style1"> 
                            <%
							if showapr3="s" and showprova3="s" then
	if pr3="" or isnull(pr3) then
	else							
	pr3 = formatNumber(pr3,1)												
							response.Write("&nbsp;"&pr3)
	end if								

							end if
							%>
                            </font></div></td>
                        <td width="30"
> <div align="center"><font class="style1"> 
                            <%
							if showapr3="s" and showprova3="s" then					
							response.Write(va_me3)
							end if
							%>
                            </font></div></td>
                        <td width="30"
> <div align="center"><font class="style1"> 
                            <%
							if showapr3="s" and showprova3="s" then					
							response.Write(va_mc3)
							end if
							%>
                            </font></div></td>
                        <td width="30"
> <div align="center"><font class="style1"> 
                            <%
							if showapr3="s" and showprova3="s" then
							response.Write(va_faltas3)
							else
							end if
							%>
                            </font></div></td>
                        <td width="30"
> <div align="center"><font class="style1"> 
<%
					if showapr3="s" and showprova3="s" then	
						if va_mc3="&nbsp;" or isnull(va_mc3) or va_mc3="" then
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
                            </font></div></td>
                        <td width="38"
> <div align="center"><font class="style1"> 
                            <%
							if resultado1="ece"then							
							if showapr4="s" and showprova4="s" then
	if va_sapr4="" or isnull(va_sapr4) then
	va_sapr4=""
	else						
	va_sapr4 = formatNumber(va_sapr4,1)
	end if												
							response.Write(va_sapr4)
							end if
							end if
							%>
                            </font></div></td>
                        <td width="38"
> <div align="center"><font class="style1"> 
                            <%
							if resultado1="ece"then
							if showapr4="s" and showprova4="s" then
	if pr4="" or isnull(pr4) then
	else							
	pr4 = formatNumber(pr4,1)								
				response.Write("&nbsp;"&pr4)
	end if
							end if
							end if
							%>
                            </font></div></td>
                        <td width="38"
> <div align="center"><font class="style1"> 
                            <%
							if resultado1="ece"then
							if showapr4="s" and showprova4="s" then					
							response.Write(va_me4)
							end if
							end if
							%>
                            </font></div></td>
                        <td width="38"
> <div align="center"><font class="style1"> 
                            <%
							if resultado1="ece" then
							if showapr4="s" and showprova4="s" then					
							response.Write(va_mc4)
							end if
							end if
							%>
                            </font></div></td>
                        <td width="30"
> <div align="center"><font class="style1"> 
                            <%
			  if resultado1="ece" then 			  
					if showapr4="s" and showprova4="s" then	
						if va_mc4="&nbsp;" or isnull(va_mc4) or va_mc4="" then
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
                            </font></div></td>
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
                      <tr> 
                        <td colspan="22" 
><div align="right"><font class="style1">* Esta nota est&aacute; sujeita a altera&ccedil;&otilde;es pela 1&ordf; 
                          Etapa Complementar de Estudos (Vide o Boletim de Avalia&ccedil;&otilde;es 
                          do 2&ordm; Trimestre).</font></div></td>
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
