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



 	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0

 	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
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


		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Controle"
		RS4.Open SQL4, CON
	
co_apr1=RS4("CO_apr1")
co_apr2=RS4("CO_apr2")
co_apr3=RS4("CO_apr3")
co_apr4=RS4("CO_apr4")
co_apr5=RS4("CO_apr5")
co_apr6=RS4("CO_apr6")
co_prova1=RS4("CO_prova1")
co_prova2=RS4("CO_prova2")
co_prova3=RS4("CO_prova3")
co_prova4=RS4("CO_prova4")
co_prova5=RS4("CO_prova5")
co_prova6=RS4("CO_prova6")	
periodo_check=periodo_check*1		
if periodo_check=1 then		
		if co_apr1="D"then
		showapr_ctrl="n"
		else 
		showapr_ctrl="s"
		end if
		if co_prova1="D"then
		showprova_ctrl="n"
		else 
		showprova_ctrl="s"
		end if
elseif periodo_check=2 then		
		if co_apr2="D"then
		showapr_ctrl="n"
		else 
		showapr_ctrl="s"
		end if
		if co_prova2="D"then
		showprova_ctrl="n"
		else 
		showprova_ctrl="s"
		end if					
elseif periodo_check=3 then		
		if co_apr3="D"then
		showapr_ctrl="n"
		else 
		showapr_ctrl="s"
		end if
		if co_prova3="D"then
		showprova_ctrl="n"
		else 
		showprova_ctrl="s"
		end if
elseif periodo_check=4 then		
		if co_apr4="D"then
		showapr_ctrl="n"
		else 
		showapr_ctrl="s"
		end if
		if co_prova4="D"then
		showprova_ctrl="n"
		else 
		showprova_ctrl="s"
		end if
elseif periodo_check=5 then		
		if co_apr5="D"then
		showapr_ctrl="n"
		else 
		showapr_ctrl="s"
		end if
		if co_prova5="D"then
		showprova_ctrl="n"
		else 
		showprova_ctrl="s"
		end if
elseif periodo_check=6 then		
		if co_apr6="D"then
		showapr_ctrl="n"
		else 
		showapr_ctrl="s"
		end if
		if co_prova6="D"then
		showprova_ctrl="n"
		else 
		showprova_ctrl="s"
		end if		
end if
	


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
<form action="index.asp?opt=1" method="post"><table width="1000" height="1039" border="0" align="center" cellpadding="0" cellspacing="0" class="tabela_principal">
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
                  <tr valign="bottom"> 
          <td height="120" colspan="3"> 
              <%call cabecalho(nivel)%>
            </td>
          </tr>
          <tr class="tabela_menu"> 
            <td width="172" height="144" rowspan="4" valign="top" class="tabela_menu"><p>&nbsp;</p>
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
>Avalia&ccedil;&otilde;es Parciais</td>
                </tr>
                <tr> 
                  <td colspan="2"> 
                    <%		Set RS_tb = Server.CreateObject("ADODB.Recordset")
		SQL_tb = "SELECT * FROM TB_Da_Aula WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso&"' AND CO_Etapa ='"& etapa&"' AND CO_Turma ='"& turma&"'"
		RS_tb.Open SQL_tb, CON2

if RS_tb.eof then
botao_impressao="n"
%>
                    <div align="center"> 
                      <%response.Write("<br><br><br><br><br><font class=style1> Não existe Boletim de Avaliações para este aluno!</font>")%>
                    </div>
                    <%
else
botao_impressao="s"
notaFIL=RS_tb("TP_Nota")

if notaFIL ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na

elseif notaFIL="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb

elseif notaFIL ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc

elseif notaFIL ="TB_NOTA_E" then
		CAMINHOn = CAMINHO_ne
		
else
		response.Write("ERRO")
end if	

obr=cod&"?"&periodo_check&"?"&notaFIL

		Set CON_N = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIRn

if notaFIL="TB_NOTA_A" then
minimo_recuperacao= 60
%>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="225" class="style3"> <div align="left">Disciplina</div></td>
                        <td width="31"  class="style3"> <div align="center">T1</div></td>
                        <td width="31"  class="style3"> <div align="center">T2</div></td>
                        <td width="31"  class="style3"> <div align="center">T3</div></td>
                        <td width="31"  class="style3"> <div align="center">T4</div></td>
                        <td width="31" class="style3"> <div align="center">MT</div></td>
<!--                        <td width="31" class="style3">&nbsp;</td>-->
                        <td width="31"  class="style3"> <div align="center">PR1</div></td>
                        <td width="31"  class="style3"> <div align="center">PR2</div></td>
                        <td width="31"  class="style3"> <div align="center">PR3</div></td>
                        <td width="31" class="style3"> <div align="center"> 
                            MP</div></td>
<!--                        <td width="31" class="style3">&nbsp;</td>-->
                        <td width="31" class="style3"> <div align="center">M1</div></td>
                        <td width="31"  class="style3"> <div align="center">Bon</div></td>
                        <td width="31"  class="style3"> <div align="center">M2</div></td>
                        <td width="31"  class="style3"> <div align="center">Rec</div></td>
                        <td width="31"  class="style3"> <div align="center">M3</div></td>
                        <td width="180" class="style3"> <div align="center">Data/Hora</div></td>
                      </tr>
<!--                      <tr>
                        <td width="225" class="style3">&nbsp;</td> 
                        <td width="31"  class="style3"> <div align="center">N</div></td>
                        <td width="31"  class="style3"> <div align="center">N</div></td>
                        <td width="31"  class="style3"> <div align="center">N</div></td>
                        <td width="31"  class="style3"> <div align="center">N</div></td>
                        <td width="31" class="style3"> 
                          <div align="center">M</div></td>
                        <td width="31" class="style3"> <div align="center">P</div></td>
                        <td width="31"  class="style3"> <div align="center">N</div></td>
                        <td width="31"  class="style3"> <div align="center">N</div></td>
                        <td width="31"  class="style3"> <div align="center">N</div></td>
                        <td width="31" class="style3"> 
                          <div align="center"> M</div></td>
                        <td width="31" class="style3"> <div align="center">P</div></td>
                        <td width="31" class="style3"> <div align="center">M</div></td>
                        <td width="31"  class="style3"> <div align="center">N</div></td>
                        <td width="31"  class="style3"> <div align="center">M</div></td>
                        <td width="31"  class="style3"> <div align="center">N</div></td>
                        <td width="31"  class="style3"> <div align="center">M</div></td>
                        <td width="191" class="style3">&nbsp;</td>
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
					
else
		va_pt=RS3("PE_Teste")
		va_pp=RS3("PE_Prova")
		va_t1=RS3("VA_Teste1")
		va_t2=RS3("VA_Teste2")
		va_t3=RS3("VA_Teste3")
		va_t4=RS3("VA_Teste4")
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
		
end if

if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then										
	showapr="n"
	showprova="n"
else
		showapr=showapr_ctrl
		showprova=showprova_ctrl
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


'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
'data_inicio=""
'va_faltas=""
'		end if


%>
                      <tr class="<%response.Write(cor)%>"> 
                        <td width="225"> 
                          <%response.Write(no_materia)%>
                        </td>
                        <td width="31"> <div align="center"> 
                            <%
							if showapr="s" then							
							response.Write(va_t1)
							End IF							
							%>
                          </div></td>
                        <td width="31"> <div align="center"> 
                            <%
							if showapr="s" then					
							response.Write(va_t2)
							end if
							%>
                          </div></td>
                        <td width="31"
> <div align="center"> 
                            <%
							if showapr="s" then					
							response.Write(va_t3)
							end if
							%>
                          </div></td>
                        <td width="31"
> <div align="center"> 
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
<!--                        <td width="31"
> <div align="center"> 
                            <%
							if showapr="s" then					
							'response.Write(va_pt)
							end if
							%>
                          </div></td>-->
                        <td width="31"
> <div align="center"> 
                            <%
							if showprova="s" then					
							response.Write(va_p1)
							end if
							%>
                          </div></td>
                        <td width="31"
> <div align="center"> 
                            <%
							if showprova="s" then					
							response.Write(va_p2)
							end if
							%>
                          </div></td>
                        <td width="31"
> <div align="center"> 
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
<!--                        <td width="31"
> <div align="center"> 
                            <%
							if showprova="s" then					
							'response.Write(va_pp)
							end if
							%>
                          </div></td>-->
                        <td width="31"
> <div align="center"> 
                            <%
							if showprova="s" AND showapr="s" then					
							response.Write(va_m1)
							end if
							%>
                          </div></td>
                        <td width="31"
> <div align="center"> 
                            <%
							if showprova="s" AND showapr="s" then					
							response.Write(va_bon)
							end if
							%>
                          </div></td>
                        <td width="31"
> <div align="center"> 
                            <%
							if showprova="s" AND showapr="s" then					
							response.Write(va_m2)
							end if
							%>
                          </div></td>
                        <td width="31"
> <div align="center"> 
                            <%
							if showprova="s" AND showapr="s" then					
							response.Write(va_rec)
							end if
							%>
                          </div></td>
                        <td width="31"
> <div align="center"> 
                            <%
							if showprova="s" AND showapr="s" then					
							response.Write(va_m3)
							end if
							%>
                          </div></td>
                        <td width="180"
> <div align="center"> 
                          <%
							if showprova="s" AND showapr="s" then							
							response.Write(data_inicio)
							End if
							%></font>
                          </td>
                      </tr>
                      <%check=check+1
RSprog.MOVENEXT
wend
%>
                      <tr valign="bottom"> 
                        <td height="20" colspan="22" 
> <div align="right"><font class="style3"> T-Teste, MT–Média dos Testes, PR-Prova, 
                        MP–Média das Provas, N-Nota e M-M&eacute;dia</font></div></td>
                      </tr>
                    </table>
                    <%
elseif notaFIL="TB_NOTA_B" or notaFIL="TB_NOTA_E" then
%>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="225" class="style3"> <div align="left">Disciplina</div></td>
                        <td width="35"  class="style3"> <div align="center">T1</div></td>
                        <td width="35"  class="style3"> <div align="center">T2</div></td>
                        <td width="35" class="style3"> <div align="center">MT</div></td>
<!--                        <td width="37" class="style3">&nbsp;</td>-->
                        <td width="35"  class="style3"> <div align="center">PR1</div></td>
                        <td width="35"  class="style3"> <div align="center">S</div></td>
                        <td width="35"  class="style3"> <div align="center">PR2</div></td>
                        <td width="35" class="style3"> <div align="center"> 
                            MP</div></td>
<!--                        <td width="37" class="style3">&nbsp;</td>-->
                        <td width="35" class="style3"> <div align="center">M1</div></td>
                        <td width="35"  class="style3"> <div align="center">Bon</div></td>
                        <td width="35"  class="style3"> <div align="center">M2</div></td>
                        <td width="35"  class="style3"> <div align="center">Rec</div></td>
                        <td width="35"  class="style3"> <div align="center">M3</div></td>
                        <td width="180" class="style3"> <div align="center">Data/Hora</div></td>
                      </tr>
<!--                      <tr>
                        <td width="225" class="style3">&nbsp;</td> 
                        <td width="36"  class="style3"> <div align="center">N</div></td>
                        <td width="36"  class="style3"> <div align="center">N</div></td>
                        <td width="37" class="style3"> 
                          <div align="center">M</div></td>
                        <td width="37" class="style3"> <div align="center">P</div></td>
                        <td width="36"  class="style3"> <div align="center">N</div></td>
                        <td width="36"  class="style3"> <div align="center">N</div></td>
                        <td width="36"  class="style3"> <div align="center">N</div></td>
                        <td width="37" class="style3"> 
                          <div align="center"> M</div></td>
                        <td width="37" class="style3"> <div align="center">P</div></td>
                        <td width="37" class="style3"> <div align="center">M</div></td>
                        <td width="36"  class="style3"> <div align="center">N</div></td>
                        <td width="37"  class="style3"> <div align="center">M</div></td>
                        <td width="36"  class="style3"> <div align="center">N</div></td>
                        <td width="37"  class="style3"> <div align="center">M</div></td>
                        <td width="191" class="style3">&nbsp;</td>
                      </tr>-->
                      <%
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
											
						else
								va_pt=RS3("PE_Teste")
								va_pp=RS3("PE_Prova")
								va_t1=RS3("VA_Teste1")
								va_t2=RS3("VA_Teste2")
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
								
						end if
if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then										
	showapr="n"
	showprova="n"
else
		showapr=showapr_ctrl
		showprova=showprova_ctrl
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
						
						
						'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
						'data_inicio=""
						'va_faltas=""
						'		end if
						
						
						%>
                      <tr class="<%response.Write(cor)%>"> 
                        <td width="225"> 
                          <%response.Write(no_materia)%>
                        </td>
                        <td width="35"> <div align="center"> 
                            <%
													if showapr="s" then							
													response.Write(va_t1)
													End IF							
													%>
                          </div></td>
                        <td width="35"> <div align="center"> 
                            <%
													if showapr="s" then					
													response.Write(va_t2)
													end if
													%>
                          </div></td>
                        <td width="35"
						> 
                          <div align="center"> 
                            <%
													if showapr="s" then					
													response.Write(va_mt)
													end if
													%>
                          </div></td>
<!--                        <td width="37"
						> <div align="center"> 
                            <%
													if showapr="s" then					
													'response.Write(va_pt)
													end if
													%>
                          </div></td>-->
                        <td width="35"
						> <div align="center"> 
                            <%
													if showprova="s" then					
													response.Write(va_p1)
													end if
													%>
                          </div></td>
                        <td width="35"
						> <div align="center"> 
                            <%
													if showprova="s" then					
													response.Write(va_p2)
													end if
													%>
                          </div></td>
                        <td width="35"
						> <div align="center"> 
                            <%
													if showprova="s" then					
													response.Write(va_p3)
													end if
													%>
                          </div></td>
                        <td width="35"
						> 
                          <div align="center"> 
                            <%
													if showprova="s" and materia<>"LP" then					
													response.Write(va_mp)
													end if
													%>
                          </div></td>
<!--                        <td width="37"
						> <div align="center"> 
                            <%
													if showprova="s" then					
													'response.Write(va_pp)
													end if
													%>
                          </div></td>-->
                        <td width="35"
						> <div align="center"> 
                            <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m1)
													end if
													%>
                          </div></td>
                        <td width="35"
						> <div align="center"> 
                            <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_bon)
													end if
													%>
                          </div></td>
                        <td width="35"
						> <div align="center"> 
                            <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m2)
													end if
													%>
                          </div></td>
                        <td width="35"
						> <div align="center"> 
                            <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_rec)
													end if
													%>
                          </div></td>
                        <td width="35"
						> <div align="center"> 
                            <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m3)
													end if
													%>
                          </div></td>
                        <td width="180"
						> <div align="center"> 
                            <%
													if showprova="s" AND showapr="s" then							
													response.Write(data_inicio)
													End if
													%>
                          </div></td>
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
									
				else
						va_pt=RS3("PE_Teste")
						va_pp=RS3("PE_Prova")
						va_t1=RS3("VA_Teste1")
						va_t2=RS3("VA_Teste2")
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
						
				end if
if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then										
	showapr="n"
	showprova="n"
else
		showapr=showapr_ctrl
		showprova=showprova_ctrl
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
				
				
				'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
				'data_inicio=""
				'va_faltas=""
				'		end if
				
				
				%>
                      <tr class="<%response.Write(cor)%>"> 
                        <td width="225"> 
                          <%response.Write(no_materia)%>
                        </td>
                        <td width="35"> <div align="center"> 
                            <%
											if showapr="s" then							
											response.Write(va_t1)
											End IF							
											%>
                          </div></td>
                        <td width="35"> <div align="center"> 
                            <%
											if showapr="s" then					
											response.Write(va_t2)
											end if
											%>
                          </div></td>
                        <td width="35"
				> 
                          <div align="center"> 
                            <%
											if showapr="s" then					
											response.Write(va_mt)
											end if
											%>
                          </div></td>
<!--                        <td width="37"
				> <div align="center"> 
                            <%
											if showapr="s" then					
										'	response.Write(va_pt)
											end if
											%>
                          </div></td>-->
                        <td width="35"
				> <div align="center"> 
                            <%
											if showprova="s" then					
											response.Write(va_p1)
											end if
											%>
                          </div></td>
                        <td width="35"
				> <div align="center"> 
                            <%
											if showprova="s" then					
											response.Write(va_p2)
											end if
											%>
                          </div></td>
                        <td width="35"
				> <div align="center"> 
                            <%
											if showprova="s" then					
											response.Write(va_p3)
											end if
											%>
                          </div></td>
                        <td width="35"
				> 
                          <div align="center"> 
                            <%
											if showprova="s" and materia<>"LP" then					
											response.Write(va_mp)
											end if
											%>
                          </div></td>
<!--                        <td width="37"
				> <div align="center"> 
                            <%
											if showprova="s" then					
											'response.Write(va_pp)
											end if
											%>
                          </div></td>-->
                        <td width="35"
				> <div align="center"> 
                            <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_m1)
											end if
											%>
                          </div></td>
                        <td width="35"
				> <div align="center"> 
                            <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_bon)
											end if
											%>
                          </div></td>
                        <td width="35"
				> <div align="center"> 
                            <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_m2)
											end if
											%>
                          </div></td>
                        <td width="35"
				> <div align="center"> 
                            <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_rec)
											end if
											%>
                          </div></td>
                        <td width="35"
				> <div align="center"> 
                            <%
											if showprova="s" AND showapr="s" then					
											response.Write(va_m3)
											end if
											%>
                          </div></td>
                        <td width="180"
				> <div align="center"> 
                            <%
											if showprova="s" AND showapr="s" then							
											response.Write(data_inicio)
											End if
											%>
                          </div></td>
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
							
								Set RSnFIL = Server.CreateObject("ADODB.Recordset")
								Set RS3 = Server.CreateObject("ADODB.Recordset")
								SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia_Principal = '"& materia &"' AND CO_Materia = '"& materia_fil &"' and NU_Periodo="&periodo_check
								Set RS3 = CON_N.Execute(SQL_N)
						if RS3.EOF then
								va_pt="&nbsp;"
								va_pp="&nbsp;"
								va_t1="&nbsp;"
								va_t2="&nbsp;"
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
											
						else
								va_pt=RS3("PE_Teste")
								va_pp=RS3("PE_Prova")
								va_t1=RS3("VA_Teste1")
								va_t2=RS3("VA_Teste2")
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
								
						end if
						if isnull(va_m1) or va_m1="" then
						va_m1=0
						end if

						if isnull(va_m2) or va_m2="" then
						va_m2=0
						end if
						
						if isnull(va_m3) or va_m3="" then
						va_m3=0
						end if												
						
							peso_acumula=peso_acumula+nu_peso_fil
							m1_ac=m1_ac+(va_m1*nu_peso_fil)	
							m2_ac=m2_ac+(va_m2*nu_peso_fil)
							m3_ac=m3_ac+(va_m3*nu_peso_fil)
																										
								
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
						
						
						'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
						'data_inicio=""
						'va_faltas=""
						'		end if
						
						
						%>
                      <tr class="<%response.Write(cor)%>"> 
                        <td width="225">&nbsp;&nbsp;&nbsp; 
                          <%response.Write(no_materia_fil)%>
                        </td>
                        <td width="35"> <div align="center"> 
                            <%
													if showapr="s" then							
													response.Write(va_t1)
													End IF							
													%>
                          </div></td>
                        <td width="35"> <div align="center"> 
                            <%
													if showapr="s" then					
													response.Write(va_t2)
													end if
													%>
                          </div></td>
                        <td width="35"
						> 
                          <div align="center"> 
                            <%
													if showapr="s" then					
													response.Write(va_mt)
													end if
													%>
                          </div></td>
<!--                        <td width="37"
						> <div align="center"> 
                            <%
													if showapr="s" then					
													'response.Write(va_pt)
													end if
													%>
                          </div></td>-->
                        <td width="35"
						> <div align="center"> 
                            <%
													if showprova="s" then					
													response.Write(va_p1)
													end if
													%>
                          </div></td>
                        <td width="35"
						> <div align="center"> 
                            <%
													if showprova="s" then					
													response.Write(va_p2)
													end if
													%>
                          </div></td>
                        <td width="35"
						> <div align="center"> 
                            <%
													if showprova="s" then					
													response.Write(va_p3)
													end if
													%>
                          </div></td>
                        <td width="35"
						> 
                          <div align="center"> 
                            <%
													if showprova="s" and materia<>"LP" then					
													response.Write(va_mp)
													end if

													%>
                          </div></td>
<!--                        <td width="37"
						> <div align="center"> 
                            <%
													if showprova="s" then					
													'response.Write(va_pp)
													end if
													%>
                          </div></td>-->
                        <td width="35"
						> <div align="center"> 
                            <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m1)
													end if
													%>
                          </div></td>
                        <td width="35"
						> <div align="center"> 
                            <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_bon)
													end if
													%>
                          </div></td>
                        <td width="35"
						> <div align="center"> 
                            <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m2)
													end if
													%>
                          </div></td>
                        <td width="35"
						> <div align="center"> 
                            <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_rec)
													end if
													%>
                          </div></td>
                        <td width="35"
						> <div align="center"> 
                            <%
													if showprova="s" AND showapr="s" then					
													response.Write(va_m3)
													end if
													%>
                          </div></td>
                        <td width="180"
						> <div align="center"> 
                            <%
													if showprova="s" AND showapr="s" then							
													response.Write(data_inicio)
													End if
													%>
                          </div></td>
                      </tr>
                      <%
				RS1a.movenext
				wend
							m1_exibe=m1_ac/peso_acumula
							m2_exibe=m2_ac/peso_acumula
							m3_exibe=m3_ac/peso_acumula
							
								decimo = m1_exibe - Int(m1_exibe)
									If decimo >= 0.5 Then
										nota_arredondada = Int(m1_exibe) + 1
										m1_exibe=nota_arredondada
									Else
										nota_arredondada = Int(m1_exibe)
										m1_exibe=nota_arredondada					
									End If
								m1_exibe= formatNumber(m1_exibe,0)
								
								decimo = m2_exibe - Int(m2_exibe)
									If decimo >= 0.5 Then
										nota_arredondada = Int(m2_exibe) + 1
										m2_exibe=nota_arredondada
									Else
										nota_arredondada = Int(m2_exibe)
										m2_exibe=nota_arredondada					
									End If
								m2_exibe= formatNumber(m2_exibe,0)
								
								decimo = m3_exibe - Int(m3_exibe)
									If decimo >= 0.5 Then
										nota_arredondada = Int(m3_exibe) + 1
										m3_exibe=nota_arredondada
									Else
										nota_arredondada = Int(m3_exibe)
										m3_exibe=nota_arredondada					
									End If
								m3_exibe= formatNumber(m3_exibe,0)																
				
				%>
                      <tr class="tb_fundo_linha_media"> 
                        <td width="225">&nbsp;&nbsp;&nbsp; M&eacute;dia</td>
                        <td width="35"> <div align="center"></div></td>
                        <td width="35"> <div align="center"> </div></td>
                        <td width="35"
				> 
                        <div align="center"> </div></td>
<!--                        <td width="37"
				> <div align="center"> </div></td>-->
                        <td width="35"
				> <div align="center"> </div></td>
                        <td width="35"
				> <div align="center"> </div></td>
                        <td width="35"
				> <div align="center"> </div></td>
                        <td width="35"
				> 
                          <div align="center"> </div></td>
<!--                        <td width="37"
				> <div align="center"> </div></td>-->
                        <td width="35"
				> <div align="center"> 
                            <%
							if showprova="s" AND showapr="s" then
							response.Write(m1_exibe)
							end if%>
                          </div></td>
                        <td width="35"
				> <div align="center"> </div></td>
                        <td width="35"
				> <div align="center"> 
                            <%if showprova="s" AND showapr="s" then
							response.Write(m2_exibe)
							end if%>
                          </div></td>
                        <td width="35"
				> <div align="center"> </div></td>
                        <td width="35"
				> <div align="center"> 
                            <%if showprova="s" AND showapr="s" then
							response.Write(m3_exibe)
							end if%>
                          </div></td>
                        <td width="180"
				> <div align="center"> </div></td>
                      </tr>
                      <%
			end if
			end if

		RSprog.MOVENEXT
		wend
		%>
                      <tr valign="bottom"> 
                        <td height="20" colspan="20" 
		> <div align="right"><font class="style3"> 
        <% if etapa=6 or etapa=7 or etapa=8 or etapa=9 then
				Response.Write("T-Teste, PR-Prova, S - Simulado, MT=(T1+T2)/2, MP=(PR1+S), M3=((MTx1)+(MPx2))/3. <br>Para a Disciplina Português PR2 = Redação e M3=((MTx1)+(MPx2)+(PR2x2))/5.")			
		else
			Response.Write("T-Teste, MT–Média dos Testes, PR-Prova, MP–Média das Provas e M-M&eacute;dia")
		End if%>
		</font></div></td>
                      </tr>
                    </table>
                    <%
elseif notaFIL="TB_NOTA_C" then
%>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="225" class="style3"> <div align="left">Disciplina</div></td>
                        <td width="32"  class="style3"> <div align="center">T1</div></td>
                        <td width="32"  class="style3"> <div align="center">T2</div></td>
                        <td width="32"  class="style3"> <div align="center">T3</div></td>
                        <td width="32"  class="style3"> <div align="center">T4</div></td>
                        <td width="32" class="style3"> <div align="center">MT</div></td>
<!--                        <td width="33" class="style3">&nbsp;</td>-->
                        <td width="32"  class="style3"> <div align="center">PR1</div></td>
                        <td width="32"  class="style3"> <div align="center">PR2</div></td>
                        <td width="32" class="style3"> <div align="center"> 
                        MP</div></td>
<!--                        <td width="33" class="style3">&nbsp;</td>-->
                        <td width="32" class="style3"> <div align="center">M1</div></td>
                        <td width="32"  class="style3"> <div align="center">Bon</div></td>
                        <td width="32"  class="style3"> <div align="center">M2</div></td>
                        <td width="32"  class="style3"> <div align="center">Rec</div></td>
                        <td width="32"  class="style3"> <div align="center">M3</div></td>
                        <td width="171" class="style3"> <div align="center">Data/Hora</div></td>
                      </tr>
<!--                      <tr>
                        <td width="225" class="style3">&nbsp;</td> 
                        <td width="33"  class="style3"> <div align="center">N</div></td>
                        <td width="33"  class="style3"> <div align="center">N</div></td>
                        <td width="33"  class="style3"> <div align="center">N</div></td>
                        <td width="33"  class="style3"> <div align="center">N</div></td>
                        <td width="33" class="style3"> 
                          <div align="center">M</div></td>
                        <td width="33" class="style3"> <div align="center">P</div></td>
                        <td width="33"  class="style3"> <div align="center">N</div></td>
                        <td width="33"  class="style3"> <div align="center">N</div></td>
                        <td width="33" class="style3"> 
                          <div align="center"> M</div></td>
                        <td width="33" class="style3"> <div align="center">P</div></td>
                        <td width="33" class="style3"> <div align="center">M</div></td>
                        <td width="33"  class="style3"> <div align="center">N</div></td>
                        <td width="33"  class="style3"> <div align="center">M</div></td>
                        <td width="33"  class="style3"> <div align="center">N</div></td>
                        <td width="33"  class="style3"> <div align="center">M</div></td>
                        <td width="191" class="style3">&nbsp;</td>
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
		va_mp="&nbsp;"
		va_m1="&nbsp;"
		va_bon="&nbsp;"
		va_m2="&nbsp;"
		va_rec="&nbsp;"
		va_m3="&nbsp;"
		data_grav="&nbsp;"
		hora_grav="&nbsp;"
					
else
		va_pt=RS3("PE_Teste")
		va_pp=RS3("PE_Prova")
		va_t1=RS3("VA_Teste1")
		va_t2=RS3("VA_Teste2")
		va_t3=RS3("VA_Teste3")
		va_t4=RS3("VA_Teste4")
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
		
end if

if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then										
	showapr="n"
	showprova="n"
else
		showapr=showapr_ctrl
		showprova=showprova_ctrl
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


'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
'data_inicio=""
'va_faltas=""
'		end if


%>
                      <tr class="<%response.Write(cor)%>"> 
                        <td width="225"> 
                          <%response.Write(no_materia)%>
                        </td>
                        <td width="32"> <div align="center"> 
                            <%
							if showapr="s" then							
							response.Write(va_t1)
							End IF							
							%>
                          </div></td>
                        <td width="32"> <div align="center"> 
                            <%
							if showapr="s" then					
							response.Write(va_t2)
							end if
							%>
                          </div></td>
                        <td width="32"
> <div align="center"> 
                            <%
							if showapr="s" then					
							response.Write(va_t3)
							end if
							%>
                          </div></td>
                        <td width="32"
> <div align="center"> 
                            <%
							if showapr="s" then					
							response.Write(va_t4)
							end if
							%>
                          </div></td>
                        <td width="32"
> 
                          <div align="center"> 
                            <%
							if showapr="s" then					
							response.Write(va_mt)
							end if
							%>
                          </div></td>
<!--                        <td width="33"
> <div align="center"> 
                            <%
							if showapr="s" then					
							'response.Write(va_pt)
							end if
							%>
                          </div></td>-->
                        <td width="32"
> <div align="center"> 
                            <%
							if showprova="s" then					
							response.Write(va_p1)
							end if
							%>
                          </div></td>
                        <td width="32"
> <div align="center"> 
                            <%
							if showprova="s" then					
							response.Write(va_p2)
							end if
							%>
                          </div></td>
                        <td width="32"
> 
                          <div align="center"> 
                            <%
							if showprova="s" then					
							response.Write(va_mp)
							end if
							%>
                          </div></td>
<!--                        <td width="33"
> <div align="center"> 
                            <%
							if showprova="s" then					
							response.Write(va_pp)
							end if
							%>
                          </div></td>-->
                        <td width="32"
> <div align="center"> 
                            <%
							if showprova="s" AND showapr="s" then					
							response.Write(va_m1)
							end if
							%>
                          </div></td>
                        <td width="32"
> <div align="center"> 
                            <%
							if showprova="s" AND showapr="s" then					
							'response.Write(va_bon)
							end if
							%>
                          </div></td>
                        <td width="32"
> <div align="center"> 
                            <%
							if showprova="s" AND showapr="s" then					
							response.Write(va_m2)
							end if
							%>
                          </div></td>
                        <td width="32"
> <div align="center"> 
                            <%
							if showprova="s" AND showapr="s" then					
							response.Write(va_rec)
							end if
							%>
                          </div></td>
                        <td width="32"
> <div align="center"> 
                            <%
							if showprova="s" AND showapr="s" then					
							response.Write(va_m3)
							end if
							%>
                          </div></td>
                        <td width="171"
> <div align="center"> 
                          <%
							if showprova="s" AND showapr="s" then							
							response.Write(data_inicio)
							End if
							%></font></div>
                        </td>
                      </tr>
                      <%check=check+1
RSprog.MOVENEXT
wend
%>
                      <tr valign="bottom"> 
                        <td height="20" colspan="21" 
> <div align="right"><font class="style3"> T-Teste, MT–Soma dos Testes, PR-Prova, MP=(PR1+PR2)/2 e M3=((MTx1)+(MPx1))/2</font></div></td>
                      </tr>
                    </table>
                    <%
end if
end if%>
                  </td>
                </tr>
                <tr>
                  <td colspan="2">
<%if botao_impressao="s" then%>				  
				  <table width="200" height="20" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td class="tb_tit"><div align="center"><a href="#" class="impressao" onClick="MM_openBrWindow('imprime.asp?obr=<%=obr%>','','menubar=yes,width=1000,height=450')">Versão 
                                  para impressão</a></div></td>
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
    <td width="1000"><img src="../../img/rodape.jpg" width="1000" height="41" /></td>
  </tr>
</table>
</form>
</body>
</html>
