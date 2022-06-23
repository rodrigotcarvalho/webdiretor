<%'On Error Resume Next%>
<!--#include file="../../inc/caminhos.asp"-->
<!--#include file="../../inc/funcoes.asp"-->
<!--#include file="../../inc/funcoes2.asp"-->
<!--#include file="../../inc/funcoes6.asp"-->
<!--#include file="../../inc/bd_grade.asp"-->
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
	Set RS4 = Server.CreateObject("ADODB.Recordset")
	SQL_WF = "SELECT * FROM TB_Autoriza_WF WHERE NU_Unidade="&unidade&" AND CO_Curso='"&curso&"' and CO_Etapa='"&etapa&"'"
	RS4.Open SQL_WF, CON
	
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
'response.Write(co_prova4&"-"&showapr4&"-"&showprova4)
		if co_apr5="D"then
		showapr5="n"
		else 
		showapr5="s"
		end if
		if co_prova5="D"then
		showprova5="n"
		else 
		showprova5="s"
		end if
	
		if co_apr6="D"then
		showapr6="n"
		else 
		showapr6="s"
		end if
		if co_prova6="D"then
		showprova6="n"
		else 
		showprova6="s"
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
                        <td width="111" height="10"> 
                          <div align="center">CURSO</div></td>
                        <td width="63" height="10"> <div align="center"> ETAPA</div></td>
                        <td width="66" height="10"> <div align="center">TURMA</div></td>
                        <td width="60" height="10"> 
                          <div align="center">CHAMADA</div></td>
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
                        <td width="111" height="10"> 
                          <div align="center"> <font class="style1"> 
                            <%response.Write(no_curso)%>
                            </font></div></td>
                        <td width="63" height="10"> <div align="center"> <font class="style1"> 
                            <%response.Write(no_etapa)%>
                            </font></div></td>
                        <td width="66" height="10"> <div align="center"> <font class="style1"> 
                            <%response.Write(turma)%>
                            </font></div></td>
                        <td width="60" height="10"> 
                          <div align="center"> <font class="style1"> 
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
botao_impressao="n"
%>
                    <div align="center"> <%response.Write("<br><br><br><br><br><font class=style1> Não existe Boletim para este aluno!</font>")%>
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
												
	elseif notaFIL ="TB_NOTA_F" then
			CAMINHOn = CAMINHO_nf
			
	elseif notaFIL ="TB_NOTA_K" then
			CAMINHOn = CAMINHO_nk				
			
	elseif notaFIL ="TB_NOTA_V" then
			CAMINHOn = CAMINHO_nv									
	else
			response.Write("ERRO")
	end if
	
vetor_temp_aluno="&nbsp;"
conta_resultados=0
qtd_rec=0
libera_resultado="s"
if notaFIL="TB_NOTA_A" or notaFIL="TB_NOTA_C" then			
	%>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="252" rowspan="2" class="style3"> 
                          <div align="left"><strong>Disciplina</strong></div></td>
                        <td width="748" colspan="11" class="style3"> <div align="center"></div>
                          <div align="center">Aproveitamento</div></td>
                      </tr>
                      <tr> 
                        <td width="68" class="style3"> 
                          <div align="center">BIM 
                            1</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">BIM 
                            2</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">BIM 
                            3</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">BIM 
                            4</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">M&eacute;dia 
                            Anual</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">Result</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">Prova 
                            Final</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">M&eacute;dia 
                            Final</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">Result</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">Prova 
                            Recup</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">Result</div></td>
                      </tr>
                      <%
			rec_lancado="sim"
		
			Set RSprog = Server.CreateObject("ADODB.Recordset")
			SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
			RSprog.Open SQLprog, CON0
		
			check=2

			while not RSprog.EOF
				res_temp_disciplina="&nbsp;"
				dividendo1=0
				divisor1=0
				dividendo2=0
				divisor2=0
				dividendo3=0
				divisor3=0
				dividendo4=0
				divisor4=0
				dividendo_ma=0
				divisor_ma=0
				dividendo5=0
				divisor5=0
				dividendo_mf=0
				divisor_mf=0
				dividendo6=0
				divisor6=0
				dividendo_rec=0
				divisor_rec=0
				res1="&nbsp;"
				res2="&nbsp;"
				res3="&nbsp;"
				m2="&nbsp;"
				m3="&nbsp;"										
			
				materia=RSprog("CO_Materia")
				mae=RSprog("IN_MAE")
				fil=RSprog("IN_FIL")
				in_co=RSprog("IN_CO")
				nu_peso=RSprog("NU_Peso")
				ordem=RSprog("NU_Ordem_Boletim")
			
				Set RS1a = Server.CreateObject("ADODB.Recordset")
				SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"' order by NU_Ordem_Boletim"
				RS1a.Open SQL1a, CON0
					
				no_materia=RS1a("NO_Materia")
			
				if check mod 2 =0 then
				cor = "tb_fundo_linha_par" 
				cor2 = "tb_fundo_linha_impar" 				
				else 
				cor ="tb_fundo_linha_impar"
				cor2 = "tb_fundo_linha_par" 
				end if
			
					
					
				Set CON_N = Server.CreateObject("ADODB.Connection") 
				ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
				CON_N.Open ABRIRn
			
				for periodofil=1 to 6
				
										'Set RS_media_turma = Server.CreateObject("ADODB.Recordset")
					'SQL_media_turma = "Select AVG(VA_Media3) as media_turma from "& notaFIL &" WHERE CO_Matricula in ("& alunos_turma &") AND CO_Materia = '"& materia &"' AND NU_Periodo="&periodofil
					'Set RS_media_turma = CON_N.Execute(SQL_media_turma)
					
				'		if periodofil=1 then
				'		va_mt1=RS_media_turma("media_turma")
				'		elseif periodofil=2 then
				'		va_mt2=RS_media_turma("media_turma")
				'		elseif periodofil=3 then
				'		va_mt3=RS_media_turma("media_turma")
				'		elseif periodofil=4 then
				'		va_mt4=RS_media_turma("media_turma")
				'		elseif periodofil=5 then
				'		va_mt5=RS_media_turma("media_turma")
				'		elseif periodofil=6 then
				'		va_mt6=RS_media_turma("media_turma")
				'		end if				
				
				
						
					Set RSnFIL = Server.CreateObject("ADODB.Recordset")
					Set RS3 = Server.CreateObject("ADODB.Recordset")
					SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' AND NU_Periodo="&periodofil
					Set RS3 = CON_N.Execute(SQL_N)
				
				
				
						if RS3.EOF then
						if periodofil=1 then
						va_m31="&nbsp;"
						va_m31_exibe="&nbsp;"
						elseif periodofil=2 then
						va_m32="&nbsp;"
						va_m32_exibe="&nbsp;"
						elseif periodofil=3 then
						va_m33="&nbsp;"
						va_m33_exibe="&nbsp;"
						elseif periodofil=4 then
						va_m34="&nbsp;"
						va_m34_exibe="&nbsp;"
						elseif periodofil=5 then
						va_m35="&nbsp;"
						va_m35_exibe="&nbsp;"
						elseif periodofil=6 then
						va_m36="&nbsp;"
						va_m36_exibe="&nbsp;"
						end if	
					else
						if periodofil=1 then
						va_m31=RS3("VA_Media3")
						va_m31_exibe=RS3("VA_Media3")
						elseif periodofil=2 then
						va_m32=RS3("VA_Media3")
						va_m32_exibe=RS3("VA_Media3")
						elseif periodofil=3 then
						va_m33=RS3("VA_Media3")
						va_m33_exibe=RS3("VA_Media3")
						elseif periodofil=4 then
						va_m34=RS3("VA_Media3")
						va_m34_exibe=RS3("VA_Media3")
						elseif periodofil=5 then
						va_m35=RS3("VA_Media3")
						va_m35_exibe=RS3("VA_Media3")
						elseif periodofil=6 then
						va_m36=RS3("VA_Media3")
						va_m36_exibe=RS3("VA_Media3")
						end if
					end if
							
							if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
							dividendo1=0
							divisor1=0
							else
							dividendo1=va_m31
							divisor1=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
																		if va_m31 > 90 then
									va_m31_exibe="E"
									elseif (va_m31 > 70) and (va_m31 <= 90) then
									va_m31_exibe="MB"
									elseif (va_m31 > 60) and (va_m31 <= 70) then							
									va_m31_exibe="B"
									elseif (va_m31 > 49) and (va_m31 <= 60) then
									va_m31_exibe="R"
									else							
									va_m31_exibe="I"
									end if													
								end if							
							end if	
							
							if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
							dividendo2=0
							divisor2=0
							else
							dividendo2=va_m32
							divisor2=1
														if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then	
														if va_m32 > 90 then
									va_m32_exibe="E"
									elseif (va_m32 > 70) and (va_m32 <= 90) then
									va_m32_exibe="MB"
									elseif (va_m32 > 60) and (va_m32 <= 70) then							
									va_m32_exibe="B"
									elseif (va_m32 > 49) and (va_m32 <= 60) then
									va_m32_exibe="R"
									else							
									va_m32_exibe="I"
								end if
						end if
							end if
							
							if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
							dividendo3=0
							divisor3=0
							else
							dividendo3=va_m33
							divisor3=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
									if va_m33 > 90 then
									va_m33_exibe="E"
									elseif (va_m33 > 70) and (va_m33 <= 90) then
									va_m33_exibe="MB"
									elseif (va_m33 > 60) and (va_m33 <= 70) then							
									va_m33_exibe="B"
									elseif (va_m33 > 49) and (va_m33 <= 60) then
									va_m33_exibe="R"
									else							
									va_m33_exibe="I"
								end if
								end if
							end if
							
							if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
							dividendo4=0
							divisor4=0
							else
							dividendo4=va_m34
							divisor4=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then					
									if va_m34 > 90 then
									va_m34_exibe="E"
									elseif (va_m34 > 70) and (va_m34 <= 90) then
									va_m34_exibe="MB"
									elseif (va_m34 > 60) and (va_m34 <= 70) then							
									va_m34_exibe="B"
									elseif (va_m34 > 49) and (va_m34 <= 60) then
									va_m34_exibe="R"
									else							
									va_m34_exibe="I"
									end if
								end if
							end if
								
					dividendo_ma=dividendo1+dividendo2+dividendo3+dividendo4
					divisor_ma=divisor1+divisor2+divisor3+divisor4
					
					'response.Write(dividendo_ma&"<<")
					
					if divisor_ma<4 then
					ma="&nbsp;"
					else
					ma=dividendo_ma/divisor_ma
					end if

					if ma="&nbsp;" then
					else
					'mf=mf/10
						decimo = ma - Int(ma)
							If decimo >= 0.5 Then
								nota_arredondada = Int(ma) + 1
								ma=nota_arredondada
							Else
								nota_arredondada = Int(ma)
								ma=nota_arredondada					
							End If
						ma = formatNumber(ma,0)
						ma=ma*1						
'						if ma>67 and ma<70 then
'							ma=70
'						end if						
						'if ma>=minimo_pf then
						'res1="APR"
						'else
						'res1="PFI"
						'end if 
					end if
					
					ma = AcrescentaBonusMediaAnual(cod, materia, ma)					
					
					'response.Write(va_m35&"<br>")														
					if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
					nota_aux_m2_1="&nbsp;"
					'dividendo5=0
					'divisor5=0
					else
					nota_aux_m2_1=va_m35
					'dividendo5=va_m35
					'divisor5=1
					end if
					
					
					if isnull(va_m36) or va_m36="&nbsp;" or va_m36="" then
					nota_aux_m3_1="&nbsp;"
					'dividendo6=0
					'divisor6=0
					else
					nota_aux_m3_1=va_m36
					'dividendo6=va_m36
					'divisor6=1
					end if
					
				NEXT
				
					if ma="&nbsp;" then
						libera_resultado="n"
					else	
										
'					call regra_aprovacao (unidade,curso,etapa,turma,divisor_ma,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2)
'							res1=Session("resultado_1")
'							res2=Session("resultado_2")
'							res3=Session("resultado_3")
'							m2=Session("M2")
'							m3=Session("M3")							
							resultados=novo_regra_aprovacao (cod, materia, curso,etapa,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,"waboletim")
							medias_resultados=split(resultados,"#!#")
							
							res1=medias_resultados(1)
							res2=medias_resultados(3)
							res3=medias_resultados(5)
							m2=medias_resultados(2)
							m3=medias_resultados(4)
																	
									
								'Se a coluna for de resultado e o resultado estiver preenchido
								'Verifica se o aluno foi aprovado pelo conselho de classe
							
								if res1<>"&nbsp;" then
									tipo_media = "MA"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res1 = modifica_result
									end if		
								end if	
								if res2<>"&nbsp;" then
									tipo_media = "RF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res2 = modifica_result
									end if																										
								end if															
								if res3<>"&nbsp;" then
									tipo_media = "MF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res3 = modifica_result
									end if	
								end if								
						
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then									
									if ma<>"&nbsp;" then
										if ma > 90 then
										ma="E"
										elseif (ma > 70) and (ma <= 90) then
										ma="MB"
										elseif (ma > 60) and (ma <= 70) then							
										ma="B"
										elseif (ma > 49) and (ma <= 60) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 > 90 then
										m2="E"
										elseif (m2 > 70) and (m2 <= 90) then
										m2="MB"
										elseif (m2 > 60) and (m2 <= 70) then							
										m2="B"
										elseif (m2 > 49) and (m2 <= 60) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if m3<>"&nbsp;" then									
										if m3 > 90 then
										m3="E"
										elseif (m3 > 70) and (m3 <= 90) then
										m3="MB"
										elseif (m3 > 60) and (m3 <= 70) then							
										m3="B"
										elseif (m3 > 49) and (m3 <= 60) then
										m3="R"
										else							
										m3="I"
										end if	
									end if						
								end if	
								if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then
									mostra_res1="s"
								else
									libera_resultado="n"
									mostra_res1="n"								
								end if
								
								if mostra_res1="s" and showapr5="s" and showprova5="s" then
									mostra_res2="s"
								else
									mostra_res2="n"								
								end if
								
								if mostra_res1="s" and mostra_res2="s" and showapr6="s" and showprova6="s" then
									mostra_res3="s"
								else
									mostra_res3="n"								
								end if								

								if ((res1 = "APR" or res1 = "APC") and mostra_res1="s") or ((res2 = "APR" or res2 = "APC") and mostra_res2="s") or ((res3 = "APR" or res3 = "APC") and mostra_res3="s") then
									if res1 = "APC" or res2 = "APC" or res3 = "APC" then
										res_temp_disciplina = "APC"									
									else
										res_temp_disciplina = "APR"
									end if	
								else
									if (res1 = "REP" and mostra_res1="s") or (res2 = "REP" and mostra_res2="s") or (res3 = "REP" and mostra_res3="s") then
										res_temp_disciplina = "REP"
									else
										if res2 = "REC" and mostra_res2="s" then
											if (res3="APR" or res3="APC" or res3="REP") and mostra_res3="s" THEN
												res_temp_disciplina = res3
											else
												res_temp_disciplina = "REC"
											end if	
										else
											if res1 = "PFI" and mostra_res1="s" then
												if (res2="APR" or res3="APC" or res2="REP") and mostra_res2="s" THEN
													res_temp_disciplina = res2
												else
													res_temp_disciplina = "PFI"
												end if	
											else
												libera_resultado="n"
												res_temp_disciplina = "&nbsp;"														
											end if											
										end if										
									end if								
								end if	
								if conta_resultados = 0 then
									vetor_temp_aluno = res_temp_disciplina
								else
									vetor_temp_aluno = vetor_temp_aluno&"#!#"&res_temp_disciplina								
								end if	 
								conta_resultados = conta_resultados+1		
							end if
			%>
                      <tr> 
                        <td width="252" class="<%response.Write(cor)%>"><%response.Write(no_materia)
						  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
						  %>
                          </td>
                       <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center">
                            <%
							if showapr1="s" and showprova1="s" then																	
								response.Write(va_m31_exibe)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr2="s" and showprova2="s" then												
								response.Write(va_m32_exibe)	
							else
								response.Write("&nbsp;")														
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr3="s" and showprova3="s" then					
								response.Write(va_m33_exibe)
							else
								response.Write("&nbsp;")										
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr4="s" and showprova4="s"  then					
								response.Write(va_m34_exibe)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then
								response.Write(ma)
							else
								response.Write("&nbsp;")	
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then																				
								response.Write(res1)		
							else
								libera_resultado="n"
								response.Write("&nbsp;")												
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr5="s" and showprova5="s" then												
								response.Write(va_m35_exibe)
							else
								response.Write("&nbsp;")								
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" then				
								response.Write(m2)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" and res1<>"APR" then					
								response.Write(res2)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr6="s" and showprova6="s" then
								response.Write(m3)
							else
								response.Write("&nbsp;")	
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" and showapr6="s" and showprova6="s" then													
								response.Write(res3)	
							else
								response.Write("&nbsp;")									
							end if

							%>
                            </div></td>
                      </tr>
                      <%
			check=check+1
			RSprog.MOVENEXT
			wend
			vetor_resultados= split(vetor_temp_aluno,"#!#")						
			for vr=0 to ubound(vetor_resultados)
				resultado=vetor_resultados(vr)
				
				if resultado="" or isnull(resultado) or resultado="&nbsp;" or resultado=" " or libera_resultado="n" then
					libera_resultado="n"
				else
					if result_temp="REP" then
					else
						if result_temp="REC" then
							if resultado="REP" then	
								result_temp=resultado
							end if			
						else
							if result_temp="PFI" then	
								if resultado="REP" or resultado="REC" then	
									result_temp=resultado
								end if					
							else	
								result_temp=resultado
							end if
						end if	
						if resultado="REC" then
							qtd_rec = qtd_rec+1
						end if						
					end if					
				End if										
			Next
			curso=curso*1
			etapa=etapa*1
			if curso = 1 and etapa<6 then
				if qtd_rec>=3 then
					resultado_aluno="REP"
				else
					resultado_aluno=result_temp			
				end if	
			elseif curso = 1 and etapa>5 then
				if qtd_rec>=4 then
					resultado_aluno="REP"
				else
					resultado_aluno=result_temp			
				end if				
			else
				resultado_aluno=result_temp					
			end if	
			
			Set RSF = Server.CreateObject("ADODB.Recordset")
			SQL_N = "Select * from TB_Frequencia_Periodo WHERE CO_Matricula = "& cod
			Set RSF = CON_N.Execute(SQL_N)
			
			if RSF.eof THEN
			f1="&nbsp;"
			f2="&nbsp;"
			f3="&nbsp;"
			f4="&nbsp;"			
			else	
			f1=RSF("NU_Faltas_P1")
			f2=RSF("NU_Faltas_P2")
			f3=RSF("NU_Faltas_P3")
			f4=RSF("NU_Faltas_P4")		
			END IF		
			%>
                      <tr valign="bottom"> 
                        <td height="20" colspan="12"> <div align="right"> 
                            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                              <tr> 
                                <td width="250" valign="baseline"><font class="style1">Freq&uuml;&ecirc;ncia 
                                  (Faltas):</font></td>
                                <td width="70" valign="baseline"><div align="right"><font class="style1">Bimestre 
                                    1:</font></div></td>
                                <td width="30" valign="baseline"><font class="style1"> 
                                  <%response.Write(f1)%></font>
                                  </td>
                                <td width="70" valign="baseline"><div align="right"><font class="style1">Bimestre 
                                    2:</font></div></td>
                                <td width="30" valign="baseline"><font class="style1"> 
                                  <%response.Write(f2)%></font>
                                  </td>
                                <td width="70" valign="baseline"><div align="right"><font class="style1">Bimestre 
                                    3:</font></div></td>
                                <td width="30" valign="baseline"><font class="style1"> 
                                  <%response.Write(f3)%></font>
                                  </td>
                                <td width="70" valign="baseline"><div align="right"><font class="style1">Bimestre 
                                    4:</font></div></td>
                                <td width="30" valign="baseline"><font class="style1"> 
                                  <%response.Write(f4)%></font>
                                  </td>
                                <td width="450" valign="baseline">&nbsp; </td>
                              </tr>
                            </table>
                          </div></td>
                      </tr>
                    </table>
						<%
	elseif notaFIL="TB_NOTA_B" or notaFIL="TB_NOTA_E" or notaFIL="TB_NOTA_F" then
	%>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="252" rowspan="2" class="style3"> 
                          <div align="left"><strong>Disciplina</strong></div></td>
                        <td width="748" colspan="11" class="style3"> <div align="center"></div>
                          <div align="center">Aproveitamento</div></td>
                      </tr>
                      <tr> 
                        <td width="68" class="style3"> 
                          <div align="center">BIM 
                            1</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">BIM 
                            2</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">BIM 
                            3</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">BIM 
                            4</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">M&eacute;dia 
                            Anual</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">Result</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">Prova 
                            Final</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">M&eacute;dia 
                            Final</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">Result</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">Prova 
                            Recup</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">Result</div></td>
                      </tr>
                      <%
			rec_lancado="sim"
		
			Set RSprog = Server.CreateObject("ADODB.Recordset")
			SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
			RSprog.Open SQLprog, CON0
		
			check=2
				
			while not RSprog.EOF
			
					dividendo1=0
					divisor1=0
					dividendo2=0
					divisor2=0
					dividendo3=0
					divisor3=0
					dividendo4=0
					divisor4=0
					dividendo_ma=0
					divisor_ma=0
					dividendo5=0
					divisor5=0
					dividendo_mf=0
					divisor_mf=0
					dividendo6=0
					divisor6=0
					dividendo_rec=0
					divisor_rec=0
					res1="&nbsp;"
					res2="&nbsp;"
					res3="&nbsp;"
					m2="&nbsp;"
					m3="&nbsp;"										
			
				materia=RSprog("CO_Materia")
				mae=RSprog("IN_MAE")
				fil=RSprog("IN_FIL")
				in_co=RSprog("IN_CO")
				nu_peso=RSprog("NU_Peso")
				ordem=RSprog("NU_Ordem_Boletim")
				
				if mae=TRUE THEN
				
					Set RS1a = Server.CreateObject("ADODB.Recordset")
					SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia_Principal='"&materia&"' order by NU_Ordem_Boletim"
					RS1a.Open SQL1a, CON0
					
					if RS1a.EOF then
				
						Set RS1b = Server.CreateObject("ADODB.Recordset")
						SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"' order by NU_Ordem_Boletim"
						RS1b.Open SQL1b, CON0
							
						no_materia=RS1b("NO_Materia")
					
						if check mod 2 =0 then
						cor = "tb_fundo_linha_par" 
						cor2 = "tb_fundo_linha_impar" 				
						else 
						cor ="tb_fundo_linha_impar"
						cor2 = "tb_fundo_linha_par" 
						end if					
							
						Set CON_N = Server.CreateObject("ADODB.Connection") 
						ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
						CON_N.Open ABRIRn
					
						for periodofil=1 to 6
					
								
							Set RSnFIL = Server.CreateObject("ADODB.Recordset")
							Set RS3 = Server.CreateObject("ADODB.Recordset")
							SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' AND NU_Periodo="&periodofil
							Set RS3 = CON_N.Execute(SQL_N)
								
								if RS3.EOF then
									if periodofil=1 then
									va_m31="&nbsp;"
									va_m31_exibe="&nbsp;"
									elseif periodofil=2 then
									va_m32="&nbsp;"
									va_m32_exibe="&nbsp;"
									elseif periodofil=3 then
									va_m33="&nbsp;"
									va_m33_exibe="&nbsp;"
									elseif periodofil=4 then
									va_m34="&nbsp;"
									va_m34_exibe="&nbsp;"
									elseif periodofil=5 then
									va_m35="&nbsp;"
									va_m35_exibe="&nbsp;"
									elseif periodofil=6 then
									va_m36="&nbsp;"
									va_m36_exibe="&nbsp;"
									end if	
								else
									if periodofil=1 then
									va_m31=RS3("VA_Media3")
									va_m31_exibe=RS3("VA_Media3")
									elseif periodofil=2 then
									va_m32=RS3("VA_Media3")
									va_m32_exibe=RS3("VA_Media3")
									elseif periodofil=3 then
									va_m33=RS3("VA_Media3")
									va_m33_exibe=RS3("VA_Media3")
									elseif periodofil=4 then
									va_m34=RS3("VA_Media3")
									va_m34_exibe=RS3("VA_Media3")
									elseif periodofil=5 then
									va_m35=RS3("VA_Media3")
									va_m35_exibe=RS3("VA_Media3")
									elseif periodofil=6 then
									va_m36=RS3("VA_Media3")
									va_m36_exibe=RS3("VA_Media3")
									end if
								end if
							
								if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
									dividendo1=0
									divisor1=0
								else
									dividendo1=va_m31
									divisor1=1
									if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
																			if va_m31 > 90 then
										va_m31_exibe="E"
										elseif (va_m31 > 70) and (va_m31 <= 90) then
										va_m31_exibe="MB"
										elseif (va_m31 > 60) and (va_m31 <= 70) then							
										va_m31_exibe="B"
										elseif (va_m31 > 49) and (va_m31 <= 60) then
										va_m31_exibe="R"
										else							
										va_m31_exibe="I"
										end if													
									end if							
								end if	
								
								if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
								dividendo2=0
								divisor2=0
								else
								dividendo2=va_m32
								divisor2=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then	
									if va_m32 > 90 then
									va_m32_exibe="E"
									elseif (va_m32 > 70) and (va_m32 <= 90) then
									va_m32_exibe="MB"
									elseif (va_m32 > 60) and (va_m32 <= 70) then							
									va_m32_exibe="B"
									elseif (va_m32 > 49) and (va_m32 <= 60) then
									va_m32_exibe="R"
									else							
									va_m32_exibe="I"
									end if
								end if
							end if
							
							if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
							dividendo3=0
							divisor3=0
							else
							dividendo3=va_m33
							divisor3=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
									if va_m33 > 90 then
									va_m33_exibe="E"
									elseif (va_m33 > 70) and (va_m33 <= 90) then
									va_m33_exibe="MB"
									elseif (va_m33 > 60) and (va_m33 <= 70) then							
									va_m33_exibe="B"
									elseif (va_m33 > 49) and (va_m33 <= 60) then
									va_m33_exibe="R"
									else							
									va_m33_exibe="I"
								end if
								end if
							end if
							
							if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
							dividendo4=0
							divisor4=0
							else
							dividendo4=va_m34
							divisor4=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then					
									if va_m34 > 90 then
									va_m34_exibe="E"
									elseif (va_m34 > 70) and (va_m34 <= 90) then
									va_m34_exibe="MB"
									elseif (va_m34 > 60) and (va_m34 <= 70) then							
									va_m34_exibe="B"
									elseif (va_m34 > 49) and (va_m34 <= 60) then
									va_m34_exibe="R"
									else							
									va_m34_exibe="I"
									end if
								end if
							end if
										
							dividendo_ma=dividendo1+dividendo2+dividendo3+dividendo4
							divisor_ma=divisor1+divisor2+divisor3+divisor4
											
							if divisor_ma<4 then
							ma="&nbsp;"
							else
							ma=dividendo_ma/divisor_ma
							end if
							
							if ma="&nbsp;" then
							else
							'mf=mf/10
								decimo = ma - Int(ma)
									If decimo >= 0.5 Then
										nota_arredondada = Int(ma) + 1
										ma=nota_arredondada
									Else
										nota_arredondada = Int(ma)
										ma=nota_arredondada					
									End If
								ma = formatNumber(ma,0)
								ma=ma*1						
'								if ma>67 and ma<70 then
'									ma=70
'								end if		
							end if
							
							ma = AcrescentaBonusMediaAnual(cod, materia, ma)							
																
							if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
							nota_aux_m2_1="&nbsp;"
		
							else
							nota_aux_m2_1=va_m35
							end if
		
						
							if isnull(va_m36) or va_m36="&nbsp;" or va_m36="" then
							nota_aux_m3_1="&nbsp;"
							else
							nota_aux_m3_1=va_m36
							end if

						NEXT
						
							if ma="&nbsp;" then
								libera_resultado="n"
							else	
												
'							call regra_aprovacao (unidade,curso,etapa,turma,divisor_ma,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2)
'							res1=Session("resultado_1")
'							res2=Session("resultado_2")
'							res3=Session("resultado_3")							
'							m2=Session("M2")
'							m3=Session("M3")	

								resultados=novo_regra_aprovacao (cod, materia, curso,etapa,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,"waboletim")
								medias_resultados=split(resultados,"#!#")
								
								res1=medias_resultados(1)
								res2=medias_resultados(3)
								res3=medias_resultados(5)
								m2=medias_resultados(2)
								m3=medias_resultados(4)
																		
								'Se a coluna for de resultado e o resultado estiver preenchido
								'Verifica se o aluno foi aprovado pelo conselho de classe
							
								if res1<>"&nbsp;" then
									tipo_media = "MA"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res1 = modifica_result
									end if		
								end if	
								if res2<>"&nbsp;" then
									tipo_media = "RF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res2 = modifica_result
									end if																										
								end if															
								if res3<>"&nbsp;" then
									tipo_media = "MF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res3 = modifica_result
									end if	
								end if						
							
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then									
									if ma<>"&nbsp;" then
										if ma > 90 then
										ma="E"
										elseif (ma > 70) and (ma <= 90) then
										ma="MB"
										elseif (ma > 60) and (ma <= 70) then							
										ma="B"
										elseif (ma > 49) and (ma <= 60) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 > 90 then
										m2="E"
										elseif (m2 > 70) and (m2 <= 90) then
										m2="MB"
										elseif (m2 > 60) and (m2 <= 70) then							
										m2="B"
										elseif (m2 > 49) and (m2 <= 60) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if m3<>"&nbsp;" then									
										if m3 > 90 then
										m3="E"
										elseif (m3 > 70) and (m3 <= 90) then
										m3="MB"
										elseif (m3 > 60) and (m3 <= 70) then							
										m3="B"
										elseif (m3 > 49) and (m3 <= 60) then
										m3="R"
										else							
										m3="I"
										end if	
									end if						
								end if							
								if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then
									mostra_res1="s"
								else
									libera_resultado="n"
									mostra_res1="n"								
								end if
								
								if mostra_res1="s" and showapr5="s" and showprova5="s" then
									mostra_res2="s"
								else
									mostra_res2="n"								
								end if
								
								if mostra_res1="s" and mostra_res2="s" and showapr6="s" and showprova6="s" then
									mostra_res3="s"
								else
									mostra_res3="n"								
								end if								
														
								if ((res1 = "APR" or res1 = "APC") and mostra_res1="s") or ((res2 = "APR" or res2 = "APC") and mostra_res2="s") or ((res3 = "APR" or res3 = "APC") and mostra_res3="s") then
									if res1 = "APC" or res2 = "APC" or res3 = "APC" then
										res_temp_disciplina = "APC"									
									else
										res_temp_disciplina = "APR"
									end if	
								else
									if (res1 = "REP" and mostra_res1="s") or (res2 = "REP" and mostra_res2="s") or (res3 = "REP" and mostra_res3="s") then
										res_temp_disciplina = "REP"
									else
										if res2 = "REC" and mostra_res2="s" then
											if (res3="APR" or res3="APC" or res3="REP") and mostra_res3="s" THEN
												res_temp_disciplina = res3
											else
												res_temp_disciplina = "REC"
											end if	
										else
											if res1 = "PFI" and mostra_res1="s" then
												if (res2="APR" or res3="APC" or res2="REP") and mostra_res2="s" THEN
													res_temp_disciplina = res2
												else
													res_temp_disciplina = "PFI"
												end if	
											else
												libera_resultado="n"
												res_temp_disciplina = "&nbsp;"														
											end if											
										end if										
									end if								
								end if		
								if conta_resultados = 0 then
									vetor_temp_aluno = res_temp_disciplina
								else
									vetor_temp_aluno = vetor_temp_aluno&"#!#"&res_temp_disciplina								
								end if	 
								conta_resultados = conta_resultados+1														
							end if
							
					%>
							  <tr> 
								<td width="252" class="<%response.Write(cor)%>"><%response.Write(no_materia)
								  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
								  %>
								  </td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center">
                            <%
							if showapr1="s" and showprova1="s" then																	
								response.Write(va_m31_exibe)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr2="s" and showprova2="s" then												
								response.Write(va_m32_exibe)	
							else
								response.Write("&nbsp;")														
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr3="s" and showprova3="s" then					
								response.Write(va_m33_exibe)
							else
								response.Write("&nbsp;")										
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr4="s" and showprova4="s"  then					
								response.Write(va_m34_exibe)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then
								response.Write(ma)
							else
								libera_resultado="n"
								response.Write("&nbsp;")	
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then																				
								response.Write(res1)		
							else
								libera_resultado="n"
								response.Write("&nbsp;")												
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr5="s" and showprova5="s" then												
								response.Write(va_m35_exibe)
							else
								response.Write("&nbsp;")								
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" then				
								response.Write(m2)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" and res1<>"APR" then					
								response.Write(res2)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr6="s" and showprova6="s" then
								response.Write(m3)
							else
								response.Write("&nbsp;")	
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" and showapr6="s" and showprova6="s" then													
								response.Write(res3)	
							else
								response.Write("&nbsp;")									
							end if

							%>
                            </div></td>
							  </tr>
							  <%
				else
						Set RS1b = Server.CreateObject("ADODB.Recordset")
						SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"' order by NU_Ordem_Boletim"
						RS1b.Open SQL1b, CON0
							
						no_materia=RS1b("NO_Materia")
					
						if check mod 2 =0 then
						cor = "tb_fundo_linha_par" 
						cor2 = "tb_fundo_linha_impar" 				
						else 
						cor ="tb_fundo_linha_impar"
						cor2 = "tb_fundo_linha_par" 
						end if					
							
						Set CON_N = Server.CreateObject("ADODB.Connection") 
						ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
						CON_N.Open ABRIRn
					
						for periodofil=1 to 6
					
								
							Set RSnFIL = Server.CreateObject("ADODB.Recordset")
							Set RS3 = Server.CreateObject("ADODB.Recordset")
							SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' AND NU_Periodo="&periodofil
							Set RS3 = CON_N.Execute(SQL_N)
								
								if RS3.EOF then
						if periodofil=1 then
						va_m31="&nbsp;"
						va_m31_exibe="&nbsp;"
						elseif periodofil=2 then
						va_m32="&nbsp;"
						va_m32_exibe="&nbsp;"
						elseif periodofil=3 then
						va_m33="&nbsp;"
						va_m33_exibe="&nbsp;"
						elseif periodofil=4 then
						va_m34="&nbsp;"
						va_m34_exibe="&nbsp;"
						elseif periodofil=5 then
						va_m35="&nbsp;"
						va_m35_exibe="&nbsp;"
						elseif periodofil=6 then
						va_m36="&nbsp;"
						va_m36_exibe="&nbsp;"
						end if	
					else
						if periodofil=1 then
						va_m31=RS3("VA_Media3")
						va_m31_exibe=RS3("VA_Media3")
						elseif periodofil=2 then
						va_m32=RS3("VA_Media3")
						va_m32_exibe=RS3("VA_Media3")
						elseif periodofil=3 then
						va_m33=RS3("VA_Media3")
						va_m33_exibe=RS3("VA_Media3")
						elseif periodofil=4 then
						va_m34=RS3("VA_Media3")
						va_m34_exibe=RS3("VA_Media3")
						elseif periodofil=5 then
						va_m35=RS3("VA_Media3")
						va_m35_exibe=RS3("VA_Media3")
						elseif periodofil=6 then
						va_m36=RS3("VA_Media3")
						va_m36_exibe=RS3("VA_Media3")
						end if
					end if
							
							if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
							dividendo1=0
							divisor1=0
							else
							dividendo1=va_m31
							divisor1=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
																		if va_m31 > 90 then
									va_m31_exibe="E"
									elseif (va_m31 > 70) and (va_m31 <= 90) then
									va_m31_exibe="MB"
									elseif (va_m31 > 60) and (va_m31 <= 70) then							
									va_m31_exibe="B"
									elseif (va_m31 > 49) and (va_m31 <= 60) then
									va_m31_exibe="R"
									else							
									va_m31_exibe="I"
									end if													
								end if							
							end if	
							
							if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
							dividendo2=0
							divisor2=0
							else
							dividendo2=va_m32
							divisor2=1
														if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then	
														if va_m32 > 90 then
									va_m32_exibe="E"
									elseif (va_m32 > 70) and (va_m32 <= 90) then
									va_m32_exibe="MB"
									elseif (va_m32 > 60) and (va_m32 <= 70) then							
									va_m32_exibe="B"
									elseif (va_m32 > 49) and (va_m32 <= 60) then
									va_m32_exibe="R"
									else							
									va_m32_exibe="I"
								end if
						end if
							end if
							
							if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
							dividendo3=0
							divisor3=0
							else
							dividendo3=va_m33
							divisor3=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
									if va_m33 > 90 then
									va_m33_exibe="E"
									elseif (va_m33 > 70) and (va_m33 <= 90) then
									va_m33_exibe="MB"
									elseif (va_m33 > 60) and (va_m33 <= 70) then							
									va_m33_exibe="B"
									elseif (va_m33 > 49) and (va_m33 <= 60) then
									va_m33_exibe="R"
									else							
									va_m33_exibe="I"
								end if
								end if
							end if
							
							if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
							dividendo4=0
							divisor4=0
							else
							dividendo4=va_m34
							divisor4=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then					
									if va_m34 > 90 then
									va_m34_exibe="E"
									elseif (va_m34 > 70) and (va_m34 <= 90) then
									va_m34_exibe="MB"
									elseif (va_m34 > 60) and (va_m34 <= 70) then							
									va_m34_exibe="B"
									elseif (va_m34 > 49) and (va_m34 <= 60) then
									va_m34_exibe="R"
									else							
									va_m34_exibe="I"
									end if
								end if
							end if
										
							dividendo_ma=dividendo1+dividendo2+dividendo3+dividendo4
							divisor_ma=divisor1+divisor2+divisor3+divisor4
											
							if divisor_ma<4 then
							ma="&nbsp;"
							else
							ma=dividendo_ma/divisor_ma
							end if
							
							if ma="&nbsp;" then
							else
							'mf=mf/10
								decimo = ma - Int(ma)
									If decimo >= 0.5 Then
										nota_arredondada = Int(ma) + 1
										ma=nota_arredondada
									Else
										nota_arredondada = Int(ma)
										ma=nota_arredondada					
									End If
								ma = formatNumber(ma,0)
								ma=ma*1						
'								if ma>67 and ma<70 then
'									ma=70
'								end if		
							end if
							
							ma = AcrescentaBonusMediaAnual(cod, materia, ma)							
																
							if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
							nota_aux_m2_1="&nbsp;"
		
							else
							nota_aux_m2_1=va_m35
							end if
		
						
							if isnull(va_m36) or va_m36="&nbsp;" or va_m36="" then
							nota_aux_m3_1="&nbsp;"
							else
							nota_aux_m3_1=va_m36
							end if
		
						NEXT
						
							if ma="&nbsp;" then
							else	
												
'							call regra_aprovacao (unidade,curso,etapa,turma,divisor_ma,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2)
'							res1=Session("resultado_1")
'							res2=Session("resultado_2")
'							res3=Session("resultado_3")
'							m2=Session("M2")
'							m3=Session("M3")	
								resultados=novo_regra_aprovacao (cod, materia,curso,etapa,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,"waboletim")
								medias_resultados=split(resultados,"#!#")
								
								res1=medias_resultados(1)
								res2=medias_resultados(3)
								res3=medias_resultados(5)
								m2=medias_resultados(2)
								m3=medias_resultados(4)
																		
										
								'Se a coluna for de resultado e o resultado estiver preenchido
								'Verifica se o aluno foi aprovado pelo conselho de classe
							
								if res1<>"&nbsp;" then
									tipo_media = "MA"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res1 = modifica_result
									end if		
								end if	
								if res2<>"&nbsp;" then
									tipo_media = "RF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res2 = modifica_result
									end if																										
								end if															
								if res3<>"&nbsp;" then
									tipo_media = "MF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res3 = modifica_result
									end if	
								end if									
						
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then									
									if ma<>"&nbsp;" then
										if ma > 90 then
										ma="E"
										elseif (ma > 70) and (ma <= 90) then
										ma="MB"
										elseif (ma > 60) and (ma <= 70) then							
										ma="B"
										elseif (ma > 49) and (ma <= 60) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 > 90 then
										m2="E"
										elseif (m2 > 70) and (m2 <= 90) then
										m2="MB"
										elseif (m2 > 60) and (m2 <= 70) then							
										m2="B"
										elseif (m2 > 49) and (m2 <= 60) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if m3<>"&nbsp;" then									
										if m3 > 90 then
										m3="E"
										elseif (m3 > 70) and (m3 <= 90) then
										m3="MB"
										elseif (m3 > 60) and (m3 <= 70) then							
										m3="B"
										elseif (m3 > 49) and (m3 <= 60) then
										m3="R"
										else							
										m3="I"
										end if	
									end if						
end if							
							
							
							end if
							
				'RESPONSE.Write(ano_letivo&"-"&ano_letivo_prog_aula)								
				ano_letivo = ano_letivo*1
				ano_letivo_prog_aula = ano_letivo_prog_aula*1	

				IF ano_letivo<ano_letivo_prog_aula THEN
					%>
							  <tr> 
								<td width="252" class="<%response.Write(cor)%>"><%response.Write(no_materia)
								  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
								  %>
								  </td>
                       <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center">
                            <%
							if showapr1="s" and showprova1="s" then																	
								response.Write(va_m31_exibe)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr2="s" and showprova2="s" then												
								response.Write(va_m32_exibe)	
							else
								response.Write("&nbsp;")														
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr3="s" and showprova3="s" then					
								response.Write(va_m33_exibe)
							else
								response.Write("&nbsp;")										
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr4="s" and showprova4="s"  then					
								response.Write(va_m34_exibe)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then
								response.Write(ma)
							else
								response.Write("&nbsp;")	
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then																				
								response.Write(res1)		
							else
								response.Write("&nbsp;")												
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr5="s" and showprova5="s" then												
								response.Write(va_m35_exibe)
							else
								response.Write("&nbsp;")								
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" then				
								response.Write(m2)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" and res1<>"APR" then					
								response.Write(res2)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr6="s" and showprova6="s" then
								response.Write(m3)
							else
								response.Write("&nbsp;")	
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" and showapr6="s" and showprova6="s" then													
								response.Write(res3)	
							else
								response.Write("&nbsp;")									
							end if

							%>
                            </div></td>
							  </tr>							  
							<%					
							END IF	' DO IF ano_letivo<ano_letivo_prog_aula THEN			

							
							
								divisor_m_acumul=0
								peso_acumula=0
								acumula_m1=0
								m31_ac=0
								m32_ac=0			
								m33_ac=0
								m34_ac=0
								m35_ac=0
								m36_ac=0
								m31_exibe=0
								m32_exibe=0
								m33_exibe=0
								m34_exibe=0
								m35_exibe=0
								m36_exibe=0								
								nu_peso_fil=0
								dividendo1=0
								dividendo2=0
								dividendo3=0
								dividendo4=0
								dividendo5=0
								dividendo6=0
								conta1=0
								conta2=0
								conta3=0
								conta4=0
								conta5=0
								conta6=0								
								conta_fil=0
								while not RS1a.EOF
								conta_fil=conta_fil+1
							
									materia_fil=RS1a("CO_Materia")
								
											Set RS1b = Server.CreateObject("ADODB.Recordset")
											SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia_fil&"' order by NU_Ordem_Boletim"
											RS1b.Open SQL1b, CON0
											
									no_materia_fil=RS1b("NO_Materia")
									
									Set RSpa = Server.CreateObject("ADODB.Recordset")
									SQLpa= "SELECT * FROM TB_Programa_Aula where CO_Materia='"&materia_fil&"' AND CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"'"
									RSpa.Open SQLpa, CON0
															
									nu_peso_fil=RSpa("NU_Peso")						
							
							for periodofil=1 to 6	
										
											Set RSnFIL = Server.CreateObject("ADODB.Recordset")
											Set RS3 = Server.CreateObject("ADODB.Recordset")
											SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia_Principal = '"& materia &"' AND CO_Materia = '"& materia_fil &"' and NU_Periodo="&periodofil
											Set RS3 = CON_N.Execute(SQL_N)
										  						  								
								if RS3.EOF then
						if periodofil=1 then
						va_m31="&nbsp;"
						va_m31_exibe="&nbsp;"
						conta1=conta1
						elseif periodofil=2 then
						va_m32="&nbsp;"
						va_m32_exibe="&nbsp;"
						conta2=conta2
						elseif periodofil=3 then
						va_m33="&nbsp;"
						va_m33_exibe="&nbsp;"
						conta3=conta3
						elseif periodofil=4 then
						va_m34="&nbsp;"
						va_m34_exibe="&nbsp;"
						conta4=conta4
						elseif periodofil=5 then
						va_m35="&nbsp;"
						va_m35_exibe="&nbsp;"
						conta5=conta5
						elseif periodofil=6 then
						va_m36="&nbsp;"
						va_m36_exibe="&nbsp;"
						conta6=conta6
						end if	
					else
						if periodofil=1 then
						va_m31=RS3("VA_Media3")
						va_m31_exibe=RS3("VA_Media3")
								if isnull(va_m31_exibe) or va_m31_exibe="" then
								conta1=conta1
								else
								conta1=conta1+1
								end if								
						elseif periodofil=2 then
						va_m32=RS3("VA_Media3")
						va_m32_exibe=RS3("VA_Media3")
								if isnull(va_m32_exibe) or va_m32_exibe="" then
								conta2=conta2
								else
								conta2=conta2+1
								end if						
						elseif periodofil=3 then
						va_m33=RS3("VA_Media3")
						va_m33_exibe=RS3("VA_Media3")
								if isnull(va_m33_exibe) or va_m33_exibe="" then
								conta3=conta3
								else
								conta3=conta3+1
								end if
						elseif periodofil=4 then
						va_m34=RS3("VA_Media3")
						va_m34_exibe=RS3("VA_Media3")
								if isnull(va_m34_exibe) or va_m34_exibe="" then
								conta4=conta4
								else
								conta4=conta4+1
								end if						
						elseif periodofil=5 then
						va_m35=RS3("VA_Media3")
						va_m35_exibe=RS3("VA_Media3")
								if isnull(va_m35_exibe) or va_m35_exibe="" then
								conta5=conta5
								else
								conta5=conta5+1
								end if						
						elseif periodofil=6 then
						va_m36=RS3("VA_Media3")
						va_m36_exibe=RS3("VA_Media3")
								if isnull(va_m36_exibe) or va_m36_exibe="" then
								conta6=conta6
								else
								conta6=conta6+1
								end if						
						end if
					end if
							
							if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
							dividendo1=0
							divisor1=0
							else
							dividendo1=va_m31
							divisor1=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
									if va_m31 > 90 then
									va_m31_exibe="E"
									elseif (va_m31 > 70) and (va_m31 <= 90) then
									va_m31_exibe="MB"
									elseif (va_m31 > 60) and (va_m31 <= 70) then							
									va_m31_exibe="B"
									elseif (va_m31 > 49) and (va_m31 <= 60) then
									va_m31_exibe="R"
									else							
									va_m31_exibe="I"
									end if													
								end if							
							end if	
							
							if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
							dividendo2=0
							divisor2=0
							else
							dividendo2=va_m32
							divisor2=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then	
									if va_m32 > 90 then
									va_m32_exibe="E"
									elseif (va_m32 > 70) and (va_m32 <= 90) then
									va_m32_exibe="MB"
									elseif (va_m32 > 60) and (va_m32 <= 70) then							
									va_m32_exibe="B"
									elseif (va_m32 > 49) and (va_m32 <= 60) then
									va_m32_exibe="R"
									else							
									va_m32_exibe="I"
								end if
						end if
							end if
							
							if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
							dividendo3=0
							divisor3=0
							else
							dividendo3=va_m33
							divisor3=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
									if va_m33 > 90 then
									va_m33_exibe="E"
									elseif (va_m33 > 70) and (va_m33 <= 90) then
									va_m33_exibe="MB"
									elseif (va_m33 > 60) and (va_m33 <= 70) then							
									va_m33_exibe="B"
									elseif (va_m33 > 49) and (va_m33 <= 60) then
									va_m33_exibe="R"
									else							
									va_m33_exibe="I"
								end if
								end if
							end if
							
							if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
							dividendo4=0
							divisor4=0
							else
							dividendo4=va_m34
							divisor4=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then					
									if va_m34 > 90 then
									va_m34_exibe="E"
									elseif (va_m34 > 70) and (va_m34 <= 90) then
									va_m34_exibe="MB"
									elseif (va_m34 > 60) and (va_m34 <= 70) then							
									va_m34_exibe="B"
									elseif (va_m34 > 49) and (va_m34 <= 60) then
									va_m34_exibe="R"
									else							
									va_m34_exibe="I"
									end if
								end if
							end if
										
							dividendo_ma=dividendo1+dividendo2+dividendo3+dividendo4
							divisor_ma=divisor1+divisor2+divisor3+divisor4
											
							if divisor_ma<4 then
							ma="&nbsp;"
							else
							ma=dividendo_ma/divisor_ma
							end if
													
							if ma="&nbsp;" then
							else
							'mf=mf/10
								decimo = ma - Int(ma)
									If decimo >= 0.5 Then
										nota_arredondada = Int(ma) + 1
										ma=nota_arredondada
									Else
										nota_arredondada = Int(ma)
										ma=nota_arredondada					
									End If
								ma = formatNumber(ma,0)
								ma=ma*1						
'								if ma>67 and ma<70 then
'									ma=70
'								end if							
							end if
							
							ma = AcrescentaBonusMediaAnual(cod, materia, ma)								
																
							if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
							nota_aux_m2_1="&nbsp;"
							dividendo5=0
							else
							nota_aux_m2_1=va_m35
							dividendo5=va_m35
							end if
		
						
							if isnull(va_m36) or va_m36="&nbsp;" or va_m36="" then
							nota_aux_m3_1="&nbsp;"
							dividendo6=0
							else
							nota_aux_m3_1=va_m36
							dividendo6=va_m36
							end if

		
						NEXT
						
							if ma="&nbsp;" then
							else	
												
'							call regra_aprovacao (unidade,curso,etapa,turma,divisor_ma,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2)
'							res1=Session("resultado_1")
'							res2=Session("resultado_2")
'							res3=Session("resultado_3")
'							m2=Session("M2")
'							m3=Session("M3")	
								resultados=novo_regra_aprovacao (cod, materia,curso,etapa,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,"waboletim")
								medias_resultados=split(resultados,"#!#")
								
								res1=medias_resultados(1)
								res2=medias_resultados(3)
								res3=medias_resultados(5)
								m2=medias_resultados(2)
								m3=medias_resultados(4)
																		
										
								'Se a coluna for de resultado e o resultado estiver preenchido
								'Verifica se o aluno foi aprovado pelo conselho de classe
							
								if res1<>"&nbsp;" then
									tipo_media = "MA"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res1 = modifica_result
									end if		
								end if	
								if res2<>"&nbsp;" then
									tipo_media = "RF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res2 = modifica_result
									end if																										
								end if															
								if res3<>"&nbsp;" then
									tipo_media = "MF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res3 = modifica_result
									end if	
								end if									
						
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then									
									if ma<>"&nbsp;" then
										if ma > 90 then
										ma="E"
										elseif (ma > 70) and (ma <= 90) then
										ma="MB"
										elseif (ma > 60) and (ma <= 70) then							
										ma="B"
										elseif (ma > 49) and (ma <= 60) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 > 90 then
										m2="E"
										elseif (m2 > 70) and (m2 <= 90) then
										m2="MB"
										elseif (m2 > 60) and (m2 <= 70) then							
										m2="B"
										elseif (m2 > 49) and (m2 <= 60) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if m3<>"&nbsp;" then									
										if m3 > 90 then
										m3="E"
										elseif (m3 > 70) and (m3 <= 90) then
										m3="MB"
										elseif (m3 > 60) and (m3 <= 70) then							
										m3="B"
										elseif (m3 > 49) and (m3 <= 60) then
										m3="R"
										else							
										m3="I"
										end if	
									end if						
end if							
							
							
							end if
							
				ano_letivo = ano_letivo*1
				ano_letivo_prog_aula = ano_letivo_prog_aula*1	

				IF ano_letivo<ano_letivo_prog_aula THEN							
					%>
							  <tr> 
								<td width="252" class="<%response.Write(cor)%>">&nbsp;&nbsp;&nbsp; 
								  <%response.Write(no_materia_fil)
								  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
								  %>
								  </td>
                       <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center">
                            <%
							if showapr1="s" and showprova1="s" then																	
								response.Write(va_m31_exibe)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr2="s" and showprova2="s" then												
								response.Write(va_m32_exibe)	
							else
								response.Write("&nbsp;")														
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr3="s" and showprova3="s" then					
								response.Write(va_m33_exibe)
							else
								response.Write("&nbsp;")										
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr4="s" and showprova4="s"  then					
								response.Write(va_m34_exibe)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then
								response.Write(ma)
							else
								response.Write("&nbsp;")	
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							'if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then																				
							'	response.Write(res1)		
							'else
								response.Write("&nbsp;")												
							'end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr5="s" and showprova5="s" then												
								response.Write(va_m35_exibe)
							else
								response.Write("&nbsp;")								
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							'if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" then				
							'	response.Write(m2)
							'else
								response.Write("&nbsp;")									
							'end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							'if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" and res1<>"APR" then					
							'	response.Write(res2)
							'else
								response.Write("&nbsp;")									
							'end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							'if showapr6="s" and showprova6="s" then
							'	response.Write(m3)
							'else
								response.Write("&nbsp;")	
							'end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
						'	if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" and showapr6="s" and showprova6="s" then													
							'	response.Write(res3)	
							'else
								response.Write("&nbsp;")									
							'end if

							%>
                            </div></td>
							  </tr>
							<%																			
						END IF 'DO IF ano_letivo<ano_letivo_prog_aula THEN	
							peso_acumula=peso_acumula+nu_peso_fil
							m31_ac=m31_ac+(dividendo1*nu_peso_fil)	
							m32_ac=m32_ac+(dividendo2*nu_peso_fil)
							m33_ac=m33_ac+(dividendo3*nu_peso_fil)
							m34_ac=m34_ac+(dividendo4*nu_peso_fil)							
							m35_ac=m35_ac+(dividendo5*nu_peso_fil)
							m36_ac=m36_ac+(dividendo6*nu_peso_fil)
							RS1a.movenext
							wend

							conta1=conta1*1
							conta2=conta2*1
							conta3=conta3*1
							conta4=conta4*1
							conta5=conta5*1
							conta6=conta6*1																																			
							if conta1=0 then
								m31_exibe="&nbsp;"							
							else
								m31_exibe=m31_ac/peso_acumula
								decimo = m31_exibe - Int(m31_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m31_exibe) + 1
									m31_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m31_exibe)
									m31_exibe=nota_arredondada					
								End If						
								m31_exibe = formatNumber(m31_exibe,0)							
							end if
													
							if conta2=0 then
								m32_exibe="&nbsp;"							
							else
								m32_exibe=m32_ac/peso_acumula
								decimo = m32_exibe - Int(m32_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m32_exibe) + 1
									m32_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m32_exibe)
									m32_exibe=nota_arredondada					
								End If						
								m32_exibe = formatNumber(m32_exibe,0)							
							end if							
							
							if conta3=0 then
								m33_exibe="&nbsp;"							
							else
								m33_exibe=m33_ac/peso_acumula
								decimo = m33_exibe - Int(m33_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m33_exibe) + 1
									m33_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m33_exibe)
									m33_exibe=nota_arredondada					
								End If						
								m33_exibe = formatNumber(m33_exibe,0)							
							end if
							
							if conta4=0 then
								m34_exibe="&nbsp;"							
							else
								m34_exibe=m34_ac/peso_acumula
								decimo = m34_exibe - Int(m34_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m34_exibe) + 1
									m34_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m34_exibe)
									m34_exibe=nota_arredondada					
								End If						
								m34_exibe = formatNumber(m34_exibe,0)							
							end if
							
							if conta5=0 then
								m35_exibe="&nbsp;"							
							else
								m35_exibe=m35_ac/peso_acumula
								decimo = m35_exibe - Int(m35_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m35_exibe) + 1
									m35_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m35_exibe)
									m35_exibe=nota_arredondada					
								End If						
								m35_exibe = formatNumber(m35_exibe,0)							
							end if																					
							
							if conta6=0 then
								m36_exibe="&nbsp;"							
							else
								m36_exibe=m36_ac/peso_acumula
								decimo = m36_exibe - Int(m36_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m36_exibe) + 1
									m36_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m36_exibe)
									m36_exibe=nota_arredondada					
								End If						
								m36_exibe = formatNumber(m36_exibe,0)							
							end if		
					
							m31_mae=m31_exibe																																				
							m32_mae=m32_exibe	
							m33_mae=m33_exibe								
							m34_mae=m34_exibe							
							
								if m31_mae="&nbsp;" then
								divisor_m_acumul=divisor_m_acumul
								else
								divisor_m_acumul=divisor_m_acumul+1
								m31_mae=m31_mae*1
								end if
		
								if m32_mae="&nbsp;" then
								divisor_m_acumul=divisor_m_acumul
								else
								divisor_m_acumul=divisor_m_acumul+1		
								m32_mae=m32_mae*1						
								end if
								
								if m33_mae="&nbsp;" then
								divisor_m_acumul=divisor_m_acumul
								else
								divisor_m_acumul=divisor_m_acumul+1	
								m33_mae=m33_mae*1							
								end if
								
								if m34_mae="&nbsp;" then
								divisor_m_acumul=divisor_m_acumul
								else
								divisor_m_acumul=divisor_m_acumul+1		
								m34_mae=m34_mae*1						
								end if	
							nota_aux_m2_1=m35_exibe
							nota_aux_m3_1=m36_exibe





										
										
							minimo_exibir=4
							'response.Write(va_m31&" - "&va_m32&" - "&va_m33&" - "&va_m34&" - "&divisor_m_acumul&"<"&minimo_exibir)								
							if divisor_m_acumul<minimo_exibir then
								m_acumul="&nbsp;"
							else
								m31_mae=m31_mae*1
								m32_mae=m32_mae*1
								m33_mae=m33_mae*1
								m34_mae=m34_mae*1
															
								dividendo_m_acumul=m31_mae+m32_mae+m33_mae+m34_mae							
								m_acumul=dividendo_m_acumul/divisor_m_acumul
							end if
							
							if m_acumul="&nbsp;" then
							else
							'mf=mf/10
								decimo = m_acumul - Int(m_acumul)
									If decimo >= 0.5 Then
										nota_arredondada = Int(m_acumul) + 1
										m_acumul=nota_arredondada
									Else
										nota_arredondada = Int(m_acumul)
										m_acumul=nota_arredondada					
									End If
								m_acumul = formatNumber(m_acumul,0)
								m_acumul =m_acumul *1
'								if m_acumul >67 and m_acumul <70 then
'									m_acumul =70
'								end if	
		
							end if		
							
							m_acumul = AcrescentaBonusMediaAnual(cod, materia, m_acumul)												
							
							if m_acumul="&nbsp;" then
								libera_resultado="n"
							else	
												
'							call regra_aprovacao (unidade,curso,etapa,turma,divisor_ma,m_acumul,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2)
'							res1=Session("resultado_1")
'							res2=Session("resultado_2")
'							res3=Session("resultado_3")
'							m2=Session("M2")
'							m3=Session("M3")	
								resultados=novo_regra_aprovacao (cod, materia,curso,etapa,m_acumul,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,"waboletim")
								medias_resultados=split(resultados,"#!#")
					
								res1=medias_resultados(1)
								res2=medias_resultados(3)
								res3=medias_resultados(5)
								m2=medias_resultados(2)
								m3=medias_resultados(4)
																				
										
								'Se a coluna for de resultado e o resultado estiver preenchido
								'Verifica se o aluno foi aprovado pelo conselho de classe
							
								if res1<>"&nbsp;" then
									tipo_media = "MA"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res1 = modifica_result
									end if		
								end if	
								if res2<>"&nbsp;" then
									tipo_media = "RF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res2 = modifica_result
									end if																										
								end if															
								if res3<>"&nbsp;" then
									tipo_media = "MF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res3 = modifica_result
									end if	
								end if									
						
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then									
									if ma<>"&nbsp;" then
										if ma > 90 then
										ma="E"
										elseif (ma > 70) and (ma <= 90) then
										ma="MB"
										elseif (ma > 60) and (ma <= 70) then							
										ma="B"
										elseif (ma > 49) and (ma <= 60) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 > 90 then
										m2="E"
										elseif (m2 > 70) and (m2 <= 90) then
										m2="MB"
										elseif (m2 > 60) and (m2 <= 70) then							
										m2="B"
										elseif (m2 > 49) and (m2 <= 60) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if m3<>"&nbsp;" then									
										if m3 > 90 then
										m3="E"
										elseif (m3 > 70) and (m3 <= 90) then
										m3="MB"
										elseif (m3 > 60) and (m3 <= 70) then							
										m3="B"
										elseif (m3 > 49) and (m3 <= 60) then
										m3="R"
										else							
										m3="I"
										end if	
									end if						
								end if							
								if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then
									mostra_res1="s"
								else
									libera_resultado="n"
									mostra_res1="n"								
								end if
								
								if mostra_res1="s" and showapr5="s" and showprova5="s" then
									mostra_res2="s"
								else
									mostra_res2="n"								
								end if
							
								if mostra_res1="s" and mostra_res2="s" and showapr6="s" and showprova6="s" then
									mostra_res3="s"
								else
									mostra_res3="n"								
								end if								
														
								if ((res1 = "APR" or res1 = "APC") and mostra_res1="s") or ((res2 = "APR" or res2 = "APC") and mostra_res2="s") or ((res3 = "APR" or res3 = "APC") and mostra_res3="s") then
									if res1 = "APC" or res2 = "APC" or res3 = "APC" then
										res_temp_disciplina = "APC"									
									else
										res_temp_disciplina = "APR"
									end if	
								else
									if (res1 = "REP" and mostra_res1="s") or (res2 = "REP" and mostra_res2="s") or (res3 = "REP" and mostra_res3="s") then
										res_temp_disciplina = "REP"
									else
										if res2 = "REC" and mostra_res2="s" then
											if (res3="APR" or res3="APC" or res3="REP") and mostra_res3="s" THEN
												res_temp_disciplina = res3
											else
												res_temp_disciplina = "REC"
											end if	
										else
											if res1 = "PFI" and mostra_res1="s" then
												if (res2="APR" or res3="APC" or res2="REP") and mostra_res2="s" THEN
													res_temp_disciplina = res2
												else
													res_temp_disciplina = "PFI"
												end if	
											else
												libera_resultado="n"
												res_temp_disciplina = "&nbsp;"														
											end if											
										end if										
									end if								
								end if	
								if conta_resultados = 0 then
									vetor_temp_aluno = res_temp_disciplina
								else
									vetor_temp_aluno = vetor_temp_aluno&"#!#"&res_temp_disciplina								
								end if	 
								conta_resultados = conta_resultados+1		
							end if	
						ano_letivo = ano_letivo*1	
						ano_letivo_prog_aula = ano_letivo_prog_aula*1
						IF ano_letivo<ano_letivo_prog_aula THEN								
							WRK_NOME_DISC = "&nbsp;&nbsp;&nbsp; M&eacute;dia"
							WRK_CLASSE = "tb_fundo_linha_media"
							WRK_TIPO = "M"
						ELSE
							WRK_NOME_DISC = no_materia
							WRK_CLASSE = COR	
							WRK_TIPO = "N"													
						END IF ' DO IF ANO LETIVO < ANO LETIVO PROG AULA							
							%>  
							<tr class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
								<td width="252" ><%RESPONSE.Write(WRK_NOME_DISC)%>
								  </td>
                       <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center">
                            <% 
							if showapr1="s" and showprova1="s" then																	
								response.Write(m31_exibe)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr2="s" and showprova2="s" then												
								response.Write(m32_exibe)	
							else
								response.Write("&nbsp;")														
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr3="s" and showprova3="s" then					
								response.Write(m33_exibe)
							else
								response.Write("&nbsp;")										
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr4="s" and showprova4="s"  then					
								response.Write(m34_exibe)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then
								response.Write(m_acumul)
							else
								response.Write("&nbsp;")	
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then																				
								response.Write(res1)		
							else
								libera_resultado="n"
								response.Write("&nbsp;")												
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr5="s" and showprova5="s" then												
								response.Write(va_m35_exibe)
							else
								response.Write("&nbsp;")								
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" then				
								response.Write(m2)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" and res1<>"APR" then					
								response.Write(res2)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr6="s" and showprova6="s" then
								response.Write(m3)
							else
								response.Write("&nbsp;")	
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" and showapr6="s" and showprova6="s" then													
								response.Write(res3)	
							else
								response.Write("&nbsp;")									
							end if

							%>
                            </div></td>
							  </tr>						  
					<%
						end if
					end if
					check=check+1
					RSprog.MOVENEXT
					wend

			vetor_resultados= split(vetor_temp_aluno,"#!#")						
			for vr=0 to ubound(vetor_resultados)
				resultado=vetor_resultados(vr)
				
				if resultado="" or isnull(resultado) or resultado="&nbsp;" or resultado=" " or libera_resultado="n" then
					libera_resultado="n"
				else
					if result_temp="REP" then
					else
						if result_temp="REC" then
							if resultado="REP" then	
								result_temp=resultado
							end if			
						else
							if result_temp="PFI" then	
								if resultado="REP" or resultado="REC" then	
									result_temp=resultado
								end if					
							else	
								result_temp=resultado
							end if
						end if	
						if resultado="REC" then
							qtd_rec = qtd_rec+1
						end if						
					end if					
				End if										
			Next
			curso=curso*1
			etapa=etapa*1

			if curso = 1 and etapa<6 then
				if qtd_rec>=3 then
					resultado_aluno="REP"
				else
					resultado_aluno=result_temp			
				end if	
			elseif curso = 1 and etapa>5 then
				if qtd_rec>=4 then
					resultado_aluno="REP"
				else
					resultado_aluno=result_temp			
				end if				
			else
				resultado_aluno=result_temp					
			end if	

				
				Set RSF = Server.CreateObject("ADODB.Recordset")
				SQL_N = "Select * from TB_Frequencia_Periodo WHERE CO_Matricula = "& cod
				Set RSF = CON_N.Execute(SQL_N)
				
				if RSF.eof THEN
				f1="&nbsp;"
				f2="&nbsp;"
				f3="&nbsp;"
				f4="&nbsp;"			
				else	
				f1=RSF("NU_Faltas_P1")
				f2=RSF("NU_Faltas_P2")
				f3=RSF("NU_Faltas_P3")
				f4=RSF("NU_Faltas_P4")		
				END IF		
				%>
						  <tr valign="bottom"> 
							<td height="20" colspan="12"> <div align="right"> 
								
                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr valign="middle"> 
                              <td width="250" height="20" valign="baseline"><font class="style1">Freq&uuml;&ecirc;ncia 
                                (Faltas):</font></td>
									
                              <td width="70" height="20" valign="baseline"> 
                                <div align="right"><font class="style1">Bimestre 
										1:</font></div></td>
									
                              <td width="30" height="20" valign="baseline"><font class="style1"> 
                                <%response.Write(f1)%>
                                </font> </td>
									
                              <td width="70" height="20" valign="baseline"> 
                                <div align="right"><font class="style1">Bimestre 
										2:</font></div></td>
									
                              <td width="30" height="20" valign="baseline"><font class="style1"> 
                                <%response.Write(f2)%>
                                </font> </td>
									
                              <td width="70" height="20" valign="baseline"> 
                                <div align="right"><font class="style1">Bimestre 
										3:</font></div></td>
									
                              <td width="30" height="20" valign="baseline"><font class="style1"> 
                                <%response.Write(f3)%>
                                </font> </td>
									
                              <td width="70" height="20" valign="baseline"> 
                                <div align="right"><font class="style1">Bimestre 
										4:</font></div></td>
									
                              <td width="30" height="20" valign="baseline"><font class="style1"> 
                                <%response.Write(f4)%>
                                </font> </td>
									
                          <td width="450" height="20" valign="baseline">&nbsp; </td>
								  </tr>
								</table>
							  </div></td>
						  </tr> 
                    </table>
<%	
	elseif notaFIL="TB_NOTA_K" then
	%>
<!--<table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="252" rowspan="2" class="style3"> 
                          <div align="left"><strong>Disciplina</strong></div></td>
                        <td width="748" colspan="11" class="style3"> <div align="center"></div>
                          <div align="center">Aproveitamento</div></td>
                      </tr>
                      <tr> 
                        <td width="68" class="style3"> 
                          <div align="center">BIM 
                            1</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">BIM 
                            2</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">BIM 
                            3</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">BIM 
                            4</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">M&eacute;dia 
                            Anual</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">Result</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">Prova 
                            Final</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">M&eacute;dia 
                            Final</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">Result</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">Prova 
                            Recup</div></td>
                        <td width="68" class="style3"> 
                          <div align="center">Result</div></td>
                      </tr>
                      <%
			rec_lancado="sim"
		
			Set RSprog = Server.CreateObject("ADODB.Recordset")
			SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
			RSprog.Open SQLprog, CON0
		
			check=2
				
			while not RSprog.EOF
			
					dividendo1=0
					divisor1=0
					dividendo2=0
					divisor2=0
					dividendo3=0
					divisor3=0
					dividendo4=0
					divisor4=0
					dividendo_ma=0
					divisor_ma=0
					dividendo5=0
					divisor5=0
					dividendo_mf=0
					divisor_mf=0
					dividendo6=0
					divisor6=0
					dividendo_rec=0
					divisor_rec=0
					res1="&nbsp;"
					res2="&nbsp;"
					res3="&nbsp;"
					m2="&nbsp;"
					m3="&nbsp;"										
			
				materia=RSprog("CO_Materia")
				mae=RSprog("IN_MAE")
				fil=RSprog("IN_FIL")
				in_co=RSprog("IN_CO")
				nu_peso=RSprog("NU_Peso")
				ordem=RSprog("NU_Ordem_Boletim")
				
				if mae=TRUE THEN
				
					Set RS1a = Server.CreateObject("ADODB.Recordset")
					SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia_Principal='"&materia&"' order by NU_Ordem_Boletim"
					RS1a.Open SQL1a, CON0
					
					if RS1a.EOF then
				
						Set RS1b = Server.CreateObject("ADODB.Recordset")
						SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"' order by NU_Ordem_Boletim"
						RS1b.Open SQL1b, CON0
							
						no_materia=RS1b("NO_Materia")
					
						if check mod 2 =0 then
						cor = "tb_fundo_linha_par" 
						cor2 = "tb_fundo_linha_impar" 				
						else 
						cor ="tb_fundo_linha_impar"
						cor2 = "tb_fundo_linha_par" 
						end if					
							
						Set CON_N = Server.CreateObject("ADODB.Connection") 
						ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
						CON_N.Open ABRIRn
					
						for periodofil=1 to 6
					
								
							Set RSnFIL = Server.CreateObject("ADODB.Recordset")
							Set RS3 = Server.CreateObject("ADODB.Recordset")
							SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' AND NU_Periodo="&periodofil				
							Set RS3 = CON_N.Execute(SQL_N)
								
								if RS3.EOF then
									if periodofil=1 then
									va_m31="&nbsp;"
									va_m31_exibe="&nbsp;"
									elseif periodofil=2 then
									va_m32="&nbsp;"
									va_m32_exibe="&nbsp;"
									elseif periodofil=3 then
									va_m33="&nbsp;"
									va_m33_exibe="&nbsp;"
									elseif periodofil=4 then
									va_m34="&nbsp;"
									va_m34_exibe="&nbsp;"
									elseif periodofil=5 then
									va_m35="&nbsp;"
									va_m35_exibe="&nbsp;"
									elseif periodofil=6 then
									va_m36="&nbsp;"
									va_m36_exibe="&nbsp;"
									end if	
								else
									if periodofil=1 then
									va_m31=RS3("VA_Media3")
									va_m31_exibe=RS3("VA_Media3")
									elseif periodofil=2 then
									va_m32=RS3("VA_Media3")
									va_m32_exibe=RS3("VA_Media3")
									elseif periodofil=3 then
									va_m33=RS3("VA_Media3")
									va_m33_exibe=RS3("VA_Media3")
									elseif periodofil=4 then
									va_m34=RS3("VA_Media3")
									va_m34_exibe=RS3("VA_Media3")
									elseif periodofil=5 then
									va_m35=RS3("VA_Media3")
									va_m35_exibe=RS3("VA_Media3")
									elseif periodofil=6 then
									va_m36=RS3("VA_Media3")
									va_m36_exibe=RS3("VA_Media3")
									end if
								end if

								if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
									dividendo1=0
									divisor1=0
								else
									dividendo1=va_m31
									divisor1=1
									if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
																			if va_m31 > 90 then
										va_m31_exibe="E"
										elseif (va_m31 > 70) and (va_m31 <= 90) then
										va_m31_exibe="MB"
										elseif (va_m31 > 60) and (va_m31 <= 70) then							
										va_m31_exibe="B"
										elseif (va_m31 > 49) and (va_m31 <= 60) then
										va_m31_exibe="R"
										else							
										va_m31_exibe="I"
										end if													
									end if							
								end if	
								
								if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
								dividendo2=0
								divisor2=0
								else
								dividendo2=va_m32
								divisor2=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then	
									if va_m32 > 90 then
									va_m32_exibe="E"
									elseif (va_m32 > 70) and (va_m32 <= 90) then
									va_m32_exibe="MB"
									elseif (va_m32 > 60) and (va_m32 <= 70) then							
									va_m32_exibe="B"
									elseif (va_m32 > 49) and (va_m32 <= 60) then
									va_m32_exibe="R"
									else							
									va_m32_exibe="I"
									end if
								end if
							end if
							
							if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
							dividendo3=0
							divisor3=0
							else
							dividendo3=va_m33
							divisor3=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
									if va_m33 > 90 then
									va_m33_exibe="E"
									elseif (va_m33 > 70) and (va_m33 <= 90) then
									va_m33_exibe="MB"
									elseif (va_m33 > 60) and (va_m33 <= 70) then							
									va_m33_exibe="B"
									elseif (va_m33 > 49) and (va_m33 <= 60) then
									va_m33_exibe="R"
									else							
									va_m33_exibe="I"
								end if
								end if
							end if
							
							if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
							dividendo4=0
							divisor4=0
							else
							dividendo4=va_m34
							divisor4=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then					
									if va_m34 > 90 then
									va_m34_exibe="E"
									elseif (va_m34 > 70) and (va_m34 <= 90) then
									va_m34_exibe="MB"
									elseif (va_m34 > 60) and (va_m34 <= 70) then							
									va_m34_exibe="B"
									elseif (va_m34 > 49) and (va_m34 <= 60) then
									va_m34_exibe="R"
									else							
									va_m34_exibe="I"
									end if
								end if
							end if
										
							dividendo_ma=dividendo1+dividendo2+dividendo3+dividendo4
							divisor_ma=divisor1+divisor2+divisor3+divisor4
											
							if divisor_ma<4 then
							ma="&nbsp;"
							else
							ma=dividendo_ma/divisor_ma
							end if
							
							if ma="&nbsp;" then
							else
							'mf=mf/10
								decimo = ma - Int(ma)
									If decimo >= 0.5 Then
										nota_arredondada = Int(ma) + 1
										ma=nota_arredondada
									Else
										nota_arredondada = Int(ma)
										ma=nota_arredondada					
									End If
								ma = formatNumber(ma,0)
								ma=ma*1						
'								if ma>67 and ma<70 then
'									ma=70
'								end if		
							end if
							
							ma = AcrescentaBonusMediaAnual(cod, materia, ma)							
																
							if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
							nota_aux_m2_1="&nbsp;"
		
							else
							nota_aux_m2_1=va_m35
							end if
		
						
							if isnull(va_m36) or va_m36="&nbsp;" or va_m36="" then
							nota_aux_m3_1="&nbsp;"
							else
							nota_aux_m3_1=va_m36
							end if

						NEXT
						
							if ma="&nbsp;" then
								libera_resultado="n"
							else	
												
'							call regra_aprovacao (unidade,curso,etapa,turma,divisor_ma,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2)
'							res1=Session("resultado_1")
'							res2=Session("resultado_2")
'							res3=Session("resultado_3")							
'							m2=Session("M2")
'							m3=Session("M3")	

								resultados=novo_regra_aprovacao (cod, materia, curso,etapa,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,"waboletim")
								medias_resultados=split(resultados,"#!#")
								
								res1=medias_resultados(1)
								res2=medias_resultados(3)
								res3=medias_resultados(5)
								m2=medias_resultados(2)
								m3=medias_resultados(4)
																		
								'Se a coluna for de resultado e o resultado estiver preenchido
								'Verifica se o aluno foi aprovado pelo conselho de classe
							
								if res1<>"&nbsp;" then
									tipo_media = "MA"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res1 = modifica_result
									end if		
								end if	
								if res2<>"&nbsp;" then
									tipo_media = "RF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res2 = modifica_result
									end if																										
								end if															
								if res3<>"&nbsp;" then
									tipo_media = "MF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res3 = modifica_result
									end if	
								end if						
							
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then									
									if ma<>"&nbsp;" then
										if ma > 90 then
										ma="E"
										elseif (ma > 70) and (ma <= 90) then
										ma="MB"
										elseif (ma > 60) and (ma <= 70) then							
										ma="B"
										elseif (ma > 49) and (ma <= 60) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 > 90 then
										m2="E"
										elseif (m2 > 70) and (m2 <= 90) then
										m2="MB"
										elseif (m2 > 60) and (m2 <= 70) then							
										m2="B"
										elseif (m2 > 49) and (m2 <= 60) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if m3<>"&nbsp;" then									
										if m3 > 90 then
										m3="E"
										elseif (m3 > 70) and (m3 <= 90) then
										m3="MB"
										elseif (m3 > 60) and (m3 <= 70) then							
										m3="B"
										elseif (m3 > 49) and (m3 <= 60) then
										m3="R"
										else							
										m3="I"
										end if	
									end if						
								end if							
								if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then
									mostra_res1="s"
								else
									libera_resultado="n"
									mostra_res1="n"								
								end if
								
								if mostra_res1="s" and showapr5="s" and showprova5="s" then
									mostra_res2="s"
								else
									mostra_res2="n"								
								end if
								
								if mostra_res1="s" and mostra_res2="s" and showapr6="s" and showprova6="s" then
									mostra_res3="s"
								else
									mostra_res3="n"								
								end if								
														
								if ((res1 = "APR" or res1 = "APC") and mostra_res1="s") or ((res2 = "APR" or res2 = "APC") and mostra_res2="s") or ((res3 = "APR" or res3 = "APC") and mostra_res3="s") then
									if res1 = "APC" or res2 = "APC" or res3 = "APC" then
										res_temp_disciplina = "APC"									
									else
										res_temp_disciplina = "APR"
									end if	
								else
									if (res1 = "REP" and mostra_res1="s") or (res2 = "REP" and mostra_res2="s") or (res3 = "REP" and mostra_res3="s") then
										res_temp_disciplina = "REP"
									else
										if res2 = "REC" and mostra_res2="s" then
											if (res3="APR" or res3="APC" or res3="REP") and mostra_res3="s" THEN
												res_temp_disciplina = res3
											else
												res_temp_disciplina = "REC"
											end if	
										else
											if res1 = "PFI" and mostra_res1="s" then
												if (res2="APR" or res3="APC" or res2="REP") and mostra_res2="s" THEN
													res_temp_disciplina = res2
												else
													res_temp_disciplina = "PFI"
												end if	
											else
												libera_resultado="n"
												res_temp_disciplina = "&nbsp;"														
											end if											
										end if										
									end if								
								end if		
								if conta_resultados = 0 then
									vetor_temp_aluno = res_temp_disciplina
								else
									vetor_temp_aluno = vetor_temp_aluno&"#!#"&res_temp_disciplina								
								end if	 
								conta_resultados = conta_resultados+1														
							end if
							
					%>
							  <tr> 
								<td width="252" class="<%response.Write(cor)%>"><%response.Write(no_materia)
								  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
								  %>
								  </td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center">
                            <%
							if showapr1="s" and showprova1="s" then																	
								response.Write(va_m31_exibe)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr2="s" and showprova2="s" then												
								response.Write(va_m32_exibe)	
							else
								response.Write("&nbsp;")														
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr3="s" and showprova3="s" then					
								response.Write(va_m33_exibe)
							else
								response.Write("&nbsp;")										
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr4="s" and showprova4="s"  then					
								response.Write(va_m34_exibe)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then
								response.Write(ma)
							else
								libera_resultado="n"
								response.Write("&nbsp;")	
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then																				
								response.Write(res1)		
							else
								libera_resultado="n"
								response.Write("&nbsp;")												
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr5="s" and showprova5="s" then												
								response.Write(va_m35_exibe)
							else
								response.Write("&nbsp;")								
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" then				
								response.Write(m2)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" and res1<>"APR" then					
								response.Write(res2)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr6="s" and showprova6="s" then
								response.Write(m3)
							else
								response.Write("&nbsp;")	
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" and showapr6="s" and showprova6="s" then													
								response.Write(res3)	
							else
								response.Write("&nbsp;")									
							end if

							%>
                            </div></td>
							  </tr>
							  <%
				else
						Set RS1b = Server.CreateObject("ADODB.Recordset")
						SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"' order by NU_Ordem_Boletim"
						RS1b.Open SQL1b, CON0
							
						no_materia=RS1b("NO_Materia")
					
						if check mod 2 =0 then
						cor = "tb_fundo_linha_par" 
						cor2 = "tb_fundo_linha_impar" 				
						else 
						cor ="tb_fundo_linha_impar"
						cor2 = "tb_fundo_linha_par" 
						end if					
							
						Set CON_N = Server.CreateObject("ADODB.Connection") 
						ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
						CON_N.Open ABRIRn
					
						for periodofil=1 to 6
					
								
							Set RSnFIL = Server.CreateObject("ADODB.Recordset")
							Set RS3 = Server.CreateObject("ADODB.Recordset")
							SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' AND NU_Periodo="&periodofil
							Set RS3 = CON_N.Execute(SQL_N)
								
								if RS3.EOF then
						if periodofil=1 then
						va_m31="&nbsp;"
						va_m31_exibe="&nbsp;"
						elseif periodofil=2 then
						va_m32="&nbsp;"
						va_m32_exibe="&nbsp;"
						elseif periodofil=3 then
						va_m33="&nbsp;"
						va_m33_exibe="&nbsp;"
						elseif periodofil=4 then
						va_m34="&nbsp;"
						va_m34_exibe="&nbsp;"
						elseif periodofil=5 then
						va_m35="&nbsp;"
						va_m35_exibe="&nbsp;"
						elseif periodofil=6 then
						va_m36="&nbsp;"
						va_m36_exibe="&nbsp;"
						end if	
					else
						if periodofil=1 then
						va_m31=RS3("VA_Media3")
						va_m31_exibe=RS3("VA_Media3")
						elseif periodofil=2 then
						va_m32=RS3("VA_Media3")
						va_m32_exibe=RS3("VA_Media3")
						elseif periodofil=3 then
						va_m33=RS3("VA_Media3")
						va_m33_exibe=RS3("VA_Media3")
						elseif periodofil=4 then
						va_m34=RS3("VA_Media3")
						va_m34_exibe=RS3("VA_Media3")
						elseif periodofil=5 then
						va_m35=RS3("VA_Media3")
						va_m35_exibe=RS3("VA_Media3")
						elseif periodofil=6 then
						va_m36=RS3("VA_Media3")
						va_m36_exibe=RS3("VA_Media3")
						end if
					end if
							
							if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
							dividendo1=0
							divisor1=0
							else
							dividendo1=va_m31
							divisor1=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
																		if va_m31 > 90 then
									va_m31_exibe="E"
									elseif (va_m31 > 70) and (va_m31 <= 90) then
									va_m31_exibe="MB"
									elseif (va_m31 > 60) and (va_m31 <= 70) then							
									va_m31_exibe="B"
									elseif (va_m31 > 49) and (va_m31 <= 60) then
									va_m31_exibe="R"
									else							
									va_m31_exibe="I"
									end if													
								end if							
							end if	
							
							if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
							dividendo2=0
							divisor2=0
							else
							dividendo2=va_m32
							divisor2=1
														if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then	
														if va_m32 > 90 then
									va_m32_exibe="E"
									elseif (va_m32 > 70) and (va_m32 <= 90) then
									va_m32_exibe="MB"
									elseif (va_m32 > 60) and (va_m32 <= 70) then							
									va_m32_exibe="B"
									elseif (va_m32 > 49) and (va_m32 <= 60) then
									va_m32_exibe="R"
									else							
									va_m32_exibe="I"
								end if
						end if
							end if
							
							if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
							dividendo3=0
							divisor3=0
							else
							dividendo3=va_m33
							divisor3=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
									if va_m33 > 90 then
									va_m33_exibe="E"
									elseif (va_m33 > 70) and (va_m33 <= 90) then
									va_m33_exibe="MB"
									elseif (va_m33 > 60) and (va_m33 <= 70) then							
									va_m33_exibe="B"
									elseif (va_m33 > 49) and (va_m33 <= 60) then
									va_m33_exibe="R"
									else							
									va_m33_exibe="I"
								end if
								end if
							end if
							
							if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
							dividendo4=0
							divisor4=0
							else
							dividendo4=va_m34
							divisor4=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then					
									if va_m34 > 90 then
									va_m34_exibe="E"
									elseif (va_m34 > 70) and (va_m34 <= 90) then
									va_m34_exibe="MB"
									elseif (va_m34 > 60) and (va_m34 <= 70) then							
									va_m34_exibe="B"
									elseif (va_m34 > 49) and (va_m34 <= 60) then
									va_m34_exibe="R"
									else							
									va_m34_exibe="I"
									end if
								end if
							end if
										
							dividendo_ma=dividendo1+dividendo2+dividendo3+dividendo4
							divisor_ma=divisor1+divisor2+divisor3+divisor4
											
							if divisor_ma<4 then
							ma="&nbsp;"
							else
							ma=dividendo_ma/divisor_ma
							end if
							
							if ma="&nbsp;" then
							else
							'mf=mf/10
								decimo = ma - Int(ma)
									If decimo >= 0.5 Then
										nota_arredondada = Int(ma) + 1
										ma=nota_arredondada
									Else
										nota_arredondada = Int(ma)
										ma=nota_arredondada					
									End If
								ma = formatNumber(ma,0)
								ma=ma*1						
'								if ma>67 and ma<70 then
'									ma=70
'								end if		
							end if
							
							ma = AcrescentaBonusMediaAnual(cod, materia, ma)							
																
							if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
							nota_aux_m2_1="&nbsp;"
		
							else
							nota_aux_m2_1=va_m35
							end if
		
						
							if isnull(va_m36) or va_m36="&nbsp;" or va_m36="" then
							nota_aux_m3_1="&nbsp;"
							else
							nota_aux_m3_1=va_m36
							end if
		
						NEXT
						
							if ma="&nbsp;" then
							else	
												
'							call regra_aprovacao (unidade,curso,etapa,turma,divisor_ma,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2)
'							res1=Session("resultado_1")
'							res2=Session("resultado_2")
'							res3=Session("resultado_3")
'							m2=Session("M2")
'							m3=Session("M3")	
								resultados=novo_regra_aprovacao (cod, materia,curso,etapa,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,"waboletim")
								medias_resultados=split(resultados,"#!#")
								
								res1=medias_resultados(1)
								res2=medias_resultados(3)
								res3=medias_resultados(5)
								m2=medias_resultados(2)
								m3=medias_resultados(4)
																		
										
								'Se a coluna for de resultado e o resultado estiver preenchido
								'Verifica se o aluno foi aprovado pelo conselho de classe
							
								if res1<>"&nbsp;" then
									tipo_media = "MA"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res1 = modifica_result
									end if		
								end if	
								if res2<>"&nbsp;" then
									tipo_media = "RF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res2 = modifica_result
									end if																										
								end if															
								if res3<>"&nbsp;" then
									tipo_media = "MF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res3 = modifica_result
									end if	
								end if									
						
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then									
									if ma<>"&nbsp;" then
										if ma > 90 then
										ma="E"
										elseif (ma > 70) and (ma <= 90) then
										ma="MB"
										elseif (ma > 60) and (ma <= 70) then							
										ma="B"
										elseif (ma > 49) and (ma <= 60) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 > 90 then
										m2="E"
										elseif (m2 > 70) and (m2 <= 90) then
										m2="MB"
										elseif (m2 > 60) and (m2 <= 70) then							
										m2="B"
										elseif (m2 > 49) and (m2 <= 60) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if m3<>"&nbsp;" then									
										if m3 > 90 then
										m3="E"
										elseif (m3 > 70) and (m3 <= 90) then
										m3="MB"
										elseif (m3 > 60) and (m3 <= 70) then							
										m3="B"
										elseif (m3 > 49) and (m3 <= 60) then
										m3="R"
										else							
										m3="I"
										end if	
									end if						
end if							
							
							
							end if
							
				'RESPONSE.Write(ano_letivo&"-"&ano_letivo_prog_aula)								
				ano_letivo = ano_letivo*1
				ano_letivo_prog_aula = ano_letivo_prog_aula*1	

				IF ano_letivo<ano_letivo_prog_aula THEN
					%>
							  <tr> 
								<td width="252" class="<%response.Write(cor)%>"><%response.Write(no_materia)
								  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
								  %>
								  </td>
                       <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center">
                            <%
							if showapr1="s" and showprova1="s" then																	
								response.Write(va_m31_exibe)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr2="s" and showprova2="s" then												
								response.Write(va_m32_exibe)	
							else
								response.Write("&nbsp;")														
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr3="s" and showprova3="s" then					
								response.Write(va_m33_exibe)
							else
								response.Write("&nbsp;")										
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr4="s" and showprova4="s"  then					
								response.Write(va_m34_exibe)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then
								response.Write(ma)
							else
								response.Write("&nbsp;")	
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then																				
								response.Write(res1)		
							else
								response.Write("&nbsp;")												
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr5="s" and showprova5="s" then												
								response.Write(va_m35_exibe)
							else
								response.Write("&nbsp;")								
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" then				
								response.Write(m2)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" and res1<>"APR" then					
								response.Write(res2)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr6="s" and showprova6="s" then
								response.Write(m3)
							else
								response.Write("&nbsp;")	
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" and showapr6="s" and showprova6="s" then													
								response.Write(res3)	
							else
								response.Write("&nbsp;")									
							end if

							%>
                            </div></td>
							  </tr>							  
							<%					
							END IF	' DO IF ano_letivo<ano_letivo_prog_aula THEN			

							
							
								divisor_m_acumul=0
								peso_acumula=0
								acumula_m1=0
								m31_ac=0
								m32_ac=0			
								m33_ac=0
								m34_ac=0
								m35_ac=0
								m36_ac=0
								m31_exibe=0
								m32_exibe=0
								m33_exibe=0
								m34_exibe=0
								m35_exibe=0
								m36_exibe=0								
								nu_peso_fil=0
								dividendo1=0
								dividendo2=0
								dividendo3=0
								dividendo4=0
								dividendo5=0
								dividendo6=0
								conta1=0
								conta2=0
								conta3=0
								conta4=0
								conta5=0
								conta6=0								
								conta_fil=0
								while not RS1a.EOF
								conta_fil=conta_fil+1
							
									materia_fil=RS1a("CO_Materia")
								
											Set RS1b = Server.CreateObject("ADODB.Recordset")
											SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia_fil&"' order by NU_Ordem_Boletim"
											RS1b.Open SQL1b, CON0
											
									no_materia_fil=RS1b("NO_Materia")
									
									Set RSpa = Server.CreateObject("ADODB.Recordset")
									SQLpa= "SELECT * FROM TB_Programa_Aula where CO_Materia='"&materia_fil&"' AND CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"'"
									RSpa.Open SQLpa, CON0
															
									nu_peso_fil=RSpa("NU_Peso")		
									
									if isnull(nu_peso_fil) or nu_peso_fil="" then
										nu_peso_fil=1
									end if				
							
							for periodofil=1 to 6	
										
											Set RSnFIL = Server.CreateObject("ADODB.Recordset")
											Set RS3 = Server.CreateObject("ADODB.Recordset")
											SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia_Principal = '"& materia &"' AND CO_Materia = '"& materia_fil &"' and NU_Periodo="&periodofil
											Set RS3 = CON_N.Execute(SQL_N)
										  						  								
								if RS3.EOF then
						if periodofil=1 then
						va_m31="&nbsp;"
						va_m31_exibe="&nbsp;"
						conta1=conta1
						elseif periodofil=2 then
						va_m32="&nbsp;"
						va_m32_exibe="&nbsp;"
						conta2=conta2
						elseif periodofil=3 then
						va_m33="&nbsp;"
						va_m33_exibe="&nbsp;"
						conta3=conta3
						elseif periodofil=4 then
						va_m34="&nbsp;"
						va_m34_exibe="&nbsp;"
						conta4=conta4
						elseif periodofil=5 then
						va_m35="&nbsp;"
						va_m35_exibe="&nbsp;"
						conta5=conta5
						elseif periodofil=6 then
						va_m36="&nbsp;"
						va_m36_exibe="&nbsp;"
						conta6=conta6
						end if	
					else
						if periodofil=1 then
						va_m31=RS3("VA_Media3")
						va_m31_exibe=RS3("VA_Media3")
								if isnull(va_m31_exibe) or va_m31_exibe="" then
								conta1=conta1
								else
								conta1=conta1+1
								end if								
						elseif periodofil=2 then
						va_m32=RS3("VA_Media3")
						va_m32_exibe=RS3("VA_Media3")
								if isnull(va_m32_exibe) or va_m32_exibe="" then
								conta2=conta2
								else
								conta2=conta2+1
								end if						
						elseif periodofil=3 then
						va_m33=RS3("VA_Media3")
						va_m33_exibe=RS3("VA_Media3")
								if isnull(va_m33_exibe) or va_m33_exibe="" then
								conta3=conta3
								else
								conta3=conta3+1
								end if

						elseif periodofil=4 then
						va_m34=RS3("VA_Media3")
						va_m34_exibe=RS3("VA_Media3")
								if isnull(va_m34_exibe) or va_m34_exibe="" then
								conta4=conta4
								else
								conta4=conta4+1
								end if						
						elseif periodofil=5 then
						va_m35=RS3("VA_Media3")
						va_m35_exibe=RS3("VA_Media3")
								if isnull(va_m35_exibe) or va_m35_exibe="" then
								conta5=conta5
								else
								conta5=conta5+1
								end if						
						elseif periodofil=6 then
						va_m36=RS3("VA_Media3")
						va_m36_exibe=RS3("VA_Media3")
								if isnull(va_m36_exibe) or va_m36_exibe="" then
								conta6=conta6
								else
								conta6=conta6+1
								end if						
						end if
					end if
							
							if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
							dividendo1=0
							divisor1=0
							else
							dividendo1=va_m31
							divisor1=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
									if va_m31 > 90 then
									va_m31_exibe="E"
									elseif (va_m31 > 70) and (va_m31 <= 90) then
									va_m31_exibe="MB"
									elseif (va_m31 > 60) and (va_m31 <= 70) then							
									va_m31_exibe="B"
									elseif (va_m31 > 49) and (va_m31 <= 60) then
									va_m31_exibe="R"
									else							
									va_m31_exibe="I"
									end if													
								end if							
							end if	
							
							if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
							dividendo2=0
							divisor2=0
							else
							dividendo2=va_m32
							divisor2=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then	
									if va_m32 > 90 then
									va_m32_exibe="E"
									elseif (va_m32 > 70) and (va_m32 <= 90) then
									va_m32_exibe="MB"
									elseif (va_m32 > 60) and (va_m32 <= 70) then							
									va_m32_exibe="B"
									elseif (va_m32 > 49) and (va_m32 <= 60) then
									va_m32_exibe="R"
									else							
									va_m32_exibe="I"
								end if
						end if
							end if
							
							if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
							dividendo3=0
							divisor3=0
							else
							dividendo3=va_m33
							divisor3=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
									if va_m33 > 90 then
									va_m33_exibe="E"
									elseif (va_m33 > 70) and (va_m33 <= 90) then
									va_m33_exibe="MB"
									elseif (va_m33 > 60) and (va_m33 <= 70) then							
									va_m33_exibe="B"
									elseif (va_m33 > 49) and (va_m33 <= 60) then
									va_m33_exibe="R"
									else							
									va_m33_exibe="I"
								end if
								end if
							end if
							
							if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
							dividendo4=0
							divisor4=0
							else
							dividendo4=va_m34
							divisor4=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then					
									if va_m34 > 90 then
									va_m34_exibe="E"
									elseif (va_m34 > 70) and (va_m34 <= 90) then
									va_m34_exibe="MB"
									elseif (va_m34 > 60) and (va_m34 <= 70) then							
									va_m34_exibe="B"
									elseif (va_m34 > 49) and (va_m34 <= 60) then
									va_m34_exibe="R"
									else							
									va_m34_exibe="I"
									end if
								end if
							end if
										
							dividendo_ma=dividendo1+dividendo2+dividendo3+dividendo4
							divisor_ma=divisor1+divisor2+divisor3+divisor4
											
							if divisor_ma<4 then
							ma="&nbsp;"
							else
							ma=dividendo_ma/divisor_ma
							end if
													
							if ma="&nbsp;" then
							else
							'mf=mf/10
								decimo = ma - Int(ma)
									If decimo >= 0.5 Then
										nota_arredondada = Int(ma) + 1
										ma=nota_arredondada
									Else
										nota_arredondada = Int(ma)
										ma=nota_arredondada					
									End If
								ma = formatNumber(ma,0)

								ma=ma*1						
'								if ma>67 and ma<70 then
'									ma=70
'								end if							
							end if
							
							ma = AcrescentaBonusMediaAnual(cod, materia, ma)								
																
							if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
							nota_aux_m2_1="&nbsp;"
							dividendo5=0
							else
							nota_aux_m2_1=va_m35
							dividendo5=va_m35
							end if
		
						
							if isnull(va_m36) or va_m36="&nbsp;" or va_m36="" then
							nota_aux_m3_1="&nbsp;"
							dividendo6=0
							else
							nota_aux_m3_1=va_m36
							dividendo6=va_m36
							end if

		
						NEXT
						
							if ma="&nbsp;" then
							else	
												
'							call regra_aprovacao (unidade,curso,etapa,turma,divisor_ma,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2)
'							res1=Session("resultado_1")
'							res2=Session("resultado_2")
'							res3=Session("resultado_3")
'							m2=Session("M2")
'							m3=Session("M3")	
								resultados=novo_regra_aprovacao (cod, materia,curso,etapa,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,"waboletim")
								medias_resultados=split(resultados,"#!#")
								
								res1=medias_resultados(1)
								res2=medias_resultados(3)
								res3=medias_resultados(5)
								m2=medias_resultados(2)
								m3=medias_resultados(4)
																		
										
								'Se a coluna for de resultado e o resultado estiver preenchido
								'Verifica se o aluno foi aprovado pelo conselho de classe
							
								if res1<>"&nbsp;" then
									tipo_media = "MA"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res1 = modifica_result
									end if		
								end if	
								if res2<>"&nbsp;" then
									tipo_media = "RF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res2 = modifica_result
									end if																										
								end if															
								if res3<>"&nbsp;" then
									tipo_media = "MF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res3 = modifica_result
									end if	
								end if									
						
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then									
									if ma<>"&nbsp;" then
										if ma > 90 then
										ma="E"
										elseif (ma > 70) and (ma <= 90) then
										ma="MB"
										elseif (ma > 60) and (ma <= 70) then							
										ma="B"
										elseif (ma > 49) and (ma <= 60) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 > 90 then
										m2="E"
										elseif (m2 > 70) and (m2 <= 90) then
										m2="MB"
										elseif (m2 > 60) and (m2 <= 70) then							
										m2="B"
										elseif (m2 > 49) and (m2 <= 60) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if m3<>"&nbsp;" then									
										if m3 > 90 then
										m3="E"
										elseif (m3 > 70) and (m3 <= 90) then
										m3="MB"
										elseif (m3 > 60) and (m3 <= 70) then							
										m3="B"
										elseif (m3 > 49) and (m3 <= 60) then
										m3="R"
										else							
										m3="I"
										end if	
									end if						
end if							
							
							
							end if
							
				ano_letivo = ano_letivo*1
				ano_letivo_prog_aula = ano_letivo_prog_aula*1	

				IF ano_letivo<ano_letivo_prog_aula THEN							
					%>
							  <tr> 
								<td width="252" class="<%response.Write(cor)%>">&nbsp;&nbsp;&nbsp; 
								  <%response.Write(no_materia_fil)
								  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
								  %>
								  </td>
                       <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center">
                            <%
							if showapr1="s" and showprova1="s" then																	
								response.Write(va_m31_exibe)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr2="s" and showprova2="s" then												
								response.Write(va_m32_exibe)	
							else
								response.Write("&nbsp;")														
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr3="s" and showprova3="s" then					
								response.Write(va_m33_exibe)
							else
								response.Write("&nbsp;")										
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr4="s" and showprova4="s"  then					
								response.Write(va_m34_exibe)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then
								response.Write(ma)
							else
								response.Write("&nbsp;")	
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							'if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then																				
							'	response.Write(res1)		
							'else
								response.Write("&nbsp;")												
							'end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							if showapr5="s" and showprova5="s" then												
								response.Write(va_m35_exibe)
							else
								response.Write("&nbsp;")								
							end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							'if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" then				
							'	response.Write(m2)
							'else
								response.Write("&nbsp;")									
							'end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							'if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" and res1<>"APR" then					
							'	response.Write(res2)
							'else
								response.Write("&nbsp;")									
							'end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
							'if showapr6="s" and showprova6="s" then
							'	response.Write(m3)
							'else
								response.Write("&nbsp;")	
							'end if
							%>
                            </div></td>
                        <td width="68" class="<%response.Write(cor)%>"> 
                          <div align="center"><%
						'	if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" and showapr6="s" and showprova6="s" then													
							'	response.Write(res3)	
							'else
								response.Write("&nbsp;")									
							'end if

							%>
                            </div></td>
							  </tr>
							<%																			
						END IF 'DO IF ano_letivo<ano_letivo_prog_aula THEN	
							'peso_acumula=peso_acumula+nu_peso_fil
							peso_acumula=nu_peso_fil							
							m31_ac=m31_ac+(dividendo1*nu_peso_fil)	
							m32_ac=m32_ac+(dividendo2*nu_peso_fil)
							m33_ac=m33_ac+(dividendo3*nu_peso_fil)
							m34_ac=m34_ac+(dividendo4*nu_peso_fil)							
							m35_ac=m35_ac+(dividendo5*nu_peso_fil)
							m36_ac=m36_ac+(dividendo6*nu_peso_fil)
							RS1a.movenext
							wend

							conta1=conta1*1
							conta2=conta2*1
							conta3=conta3*1
							conta4=conta4*1
							conta5=conta5*1
							conta6=conta6*1																																			
							if conta1=0 then
								m31_exibe="&nbsp;"							
							else
								m31_exibe=m31_ac/peso_acumula
								decimo = m31_exibe - Int(m31_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m31_exibe) + 1
									m31_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m31_exibe)
									m31_exibe=nota_arredondada					
								End If	
								if m31_exibe>100 then
									m31_exibe=100
								end if													
								m31_exibe = formatNumber(m31_exibe,0)							
							end if
							response.Write(SQL_N)
							    response.Write("<BR>"&va_m34&"-"&va_m34_exibe&"<BR>")																
							if conta2=0 then
								m32_exibe="&nbsp;"							
							else
								m32_exibe=m32_ac/peso_acumula
								decimo = m32_exibe - Int(m32_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m32_exibe) + 1
									m32_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m32_exibe)
									m32_exibe=nota_arredondada					
								End If		
								if m32_exibe>100 then
									m32_exibe=100
								end if													
								m32_exibe = formatNumber(m32_exibe,0)							
							end if							
							
							if conta3=0 then
								m33_exibe="&nbsp;"							
							else
								m33_exibe=m33_ac/peso_acumula
								decimo = m33_exibe - Int(m33_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m33_exibe) + 1
									m33_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m33_exibe)
									m33_exibe=nota_arredondada					
								End If		
								if m33_exibe>100 then
									m33_exibe=100
								end if													
								m33_exibe = formatNumber(m33_exibe,0)							
							end if
							
							if conta4=0 then
								m34_exibe="&nbsp;"							
							else
								m34_exibe=m34_ac/peso_acumula
								decimo = m34_exibe - Int(m34_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m34_exibe) + 1
									m34_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m34_exibe)
									m34_exibe=nota_arredondada					
								End If	
								if m34_exibe>100 then
									m34_exibe=100
								end if														
								m34_exibe = formatNumber(m34_exibe,0)							
							end if
							
							if conta5=0 then
								m35_exibe="&nbsp;"							
							else
								m35_exibe=m35_ac/peso_acumula
								decimo = m35_exibe - Int(m35_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m35_exibe) + 1
									m35_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m35_exibe)
									m35_exibe=nota_arredondada					
								End If		
								if m35_exibe>100 then
									m35_exibe=100
								end if													
								m35_exibe = formatNumber(m35_exibe,0)							
							end if																					
							
							if conta6=0 then
								m36_exibe="&nbsp;"							
							else
								m36_exibe=m36_ac/peso_acumula
								decimo = m36_exibe - Int(m36_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m36_exibe) + 1
									m36_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m36_exibe)
									m36_exibe=nota_arredondada					
								End If		
								if m36_exibe>100 then
									m36_exibe=100
								end if												
								m36_exibe = formatNumber(m36_exibe,0)							
							end if		
					
							m31_mae=m31_exibe																																				
							m32_mae=m32_exibe	
							m33_mae=m33_exibe								
							m34_mae=m34_exibe							
							
								if m31_mae="&nbsp;" then
								divisor_m_acumul=divisor_m_acumul
								else
								divisor_m_acumul=divisor_m_acumul+1
								m31_mae=m31_mae*1
								end if
		
								if m32_mae="&nbsp;" then
								divisor_m_acumul=divisor_m_acumul
								else
								divisor_m_acumul=divisor_m_acumul+1		
								m32_mae=m32_mae*1						
								end if
								
								if m33_mae="&nbsp;" then
								divisor_m_acumul=divisor_m_acumul
								else
								divisor_m_acumul=divisor_m_acumul+1	
								m33_mae=m33_mae*1							
								end if
								
								if m34_mae="&nbsp;" then
								divisor_m_acumul=divisor_m_acumul
								else
								divisor_m_acumul=divisor_m_acumul+1		
								m34_mae=m34_mae*1						
								end if	
							nota_aux_m2_1=m35_exibe
							nota_aux_m3_1=m36_exibe





										
										
							minimo_exibir=4
							'response.Write(va_m31&" - "&va_m32&" - "&va_m33&" - "&va_m34&" - "&divisor_m_acumul&"<"&minimo_exibir)								
							if divisor_m_acumul<minimo_exibir then
								m_acumul="&nbsp;"
							else
								m31_mae=m31_mae*1
								m32_mae=m32_mae*1
								m33_mae=m33_mae*1
								m34_mae=m34_mae*1
															
								dividendo_m_acumul=m31_mae+m32_mae+m33_mae+m34_mae							
								'm_acumul=dividendo_m_acumul/divisor_m_acumul
								m_acumul=dividendo_m_acumul
							end if
							
							if m_acumul="&nbsp;" then
							else
							'mf=mf/10
								decimo = m_acumul - Int(m_acumul)
									If decimo >= 0.5 Then
										nota_arredondada = Int(m_acumul) + 1
										m_acumul=nota_arredondada
									Else
										nota_arredondada = Int(m_acumul)
										m_acumul=nota_arredondada					
									End If
								m_acumul = formatNumber(m_acumul,0)
								m_acumul =m_acumul *1
'								if m_acumul >67 and m_acumul <70 then
'									m_acumul =70
'								end if	
		
							end if		
							
							m_acumul = AcrescentaBonusMediaAnual(cod, materia, m_acumul)												
							
							if m_acumul="&nbsp;" then
								libera_resultado="n"
							else	
												
'							call regra_aprovacao (unidade,curso,etapa,turma,divisor_ma,m_acumul,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2)
'							res1=Session("resultado_1")
'							res2=Session("resultado_2")
'							res3=Session("resultado_3")
'							m2=Session("M2")
'							m3=Session("M3")	
								resultados=novo_regra_aprovacao (cod, materia,curso,etapa,m_acumul,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,"waboletim")
								medias_resultados=split(resultados,"#!#")
					
								res1=medias_resultados(1)
								res2=medias_resultados(3)
								res3=medias_resultados(5)
								m2=medias_resultados(2)
								m3=medias_resultados(4)
																				
										
								'Se a coluna for de resultado e o resultado estiver preenchido
								'Verifica se o aluno foi aprovado pelo conselho de classe
							
								if res1<>"&nbsp;" then
									tipo_media = "MA"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res1 = modifica_result
									end if		
								end if	
								if res2<>"&nbsp;" then
									tipo_media = "RF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res2 = modifica_result
									end if																										
								end if															
								if res3<>"&nbsp;" then
									tipo_media = "MF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res3 = modifica_result
									end if	
								end if									
						
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then									
									if ma<>"&nbsp;" then
										if ma > 90 then
										ma="E"
										elseif (ma > 70) and (ma <= 90) then
										ma="MB"
										elseif (ma > 60) and (ma <= 70) then							
										ma="B"
										elseif (ma > 49) and (ma <= 60) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 > 90 then
										m2="E"
										elseif (m2 > 70) and (m2 <= 90) then
										m2="MB"
										elseif (m2 > 60) and (m2 <= 70) then							
										m2="B"
										elseif (m2 > 49) and (m2 <= 60) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if m3<>"&nbsp;" then									
										if m3 > 90 then
										m3="E"
										elseif (m3 > 70) and (m3 <= 90) then
										m3="MB"
										elseif (m3 > 60) and (m3 <= 70) then							
										m3="B"
										elseif (m3 > 49) and (m3 <= 60) then
										m3="R"
										else							
										m3="I"
										end if	
									end if						
								end if							
								if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then
									mostra_res1="s"
								else
									libera_resultado="n"
									mostra_res1="n"								
								end if
								
								if mostra_res1="s" and showapr5="s" and showprova5="s" then
									mostra_res2="s"
								else
									mostra_res2="n"								
								end if
							
								if mostra_res1="s" and mostra_res2="s" and showapr6="s" and showprova6="s" then
									mostra_res3="s"
								else
									mostra_res3="n"								
								end if								
														
								if ((res1 = "APR" or res1 = "APC") and mostra_res1="s") or ((res2 = "APR" or res2 = "APC") and mostra_res2="s") or ((res3 = "APR" or res3 = "APC") and mostra_res3="s") then
									if res1 = "APC" or res2 = "APC" or res3 = "APC" then
										res_temp_disciplina = "APC"									
									else
										res_temp_disciplina = "APR"
									end if	
								else
									if (res1 = "REP" and mostra_res1="s") or (res2 = "REP" and mostra_res2="s") or (res3 = "REP" and mostra_res3="s") then
										res_temp_disciplina = "REP"
									else
										if res2 = "REC" and mostra_res2="s" then
											if (res3="APR" or res3="APC" or res3="REP") and mostra_res3="s" THEN
												res_temp_disciplina = res3
											else
												res_temp_disciplina = "REC"
											end if	
										else
											if res1 = "PFI" and mostra_res1="s" then
												if (res2="APR" or res3="APC" or res2="REP") and mostra_res2="s" THEN
													res_temp_disciplina = res2
												else
													res_temp_disciplina = "PFI"
												end if	
											else
												libera_resultado="n"
												res_temp_disciplina = "&nbsp;"														
											end if											
										end if										
									end if								
								end if	
								if conta_resultados = 0 then
									vetor_temp_aluno = res_temp_disciplina
								else
									vetor_temp_aluno = vetor_temp_aluno&"#!#"&res_temp_disciplina								
								end if	 
								conta_resultados = conta_resultados+1		
							end if	
						ano_letivo = ano_letivo*1	
						ano_letivo_prog_aula = ano_letivo_prog_aula*1
						IF ano_letivo<ano_letivo_prog_aula THEN								
							WRK_NOME_DISC = "&nbsp;&nbsp;&nbsp; M&eacute;dia"
							WRK_CLASSE = "tb_fundo_linha_media"
							WRK_TIPO = "M"
						ELSE
							WRK_NOME_DISC = no_materia
							WRK_CLASSE = COR	
							WRK_TIPO = "N"													
						END IF ' DO IF ANO LETIVO < ANO LETIVO PROG AULA							
							%>  
							<tr class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
								<td width="252" ><%RESPONSE.Write(WRK_NOME_DISC)%>
								  </td>
                       <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center">
                            <% 
							if showapr1="s" and showprova1="s" then																	
								response.Write(m31_exibe)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr2="s" and showprova2="s" then												
								response.Write(m32_exibe)	
							else
								response.Write("&nbsp;")														
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr3="s" and showprova3="s" then					
								response.Write(m33_exibe)
							else
								response.Write("&nbsp;")										
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr4="s" and showprova4="s"  then					
								response.Write(m34_exibe)
							else
								response.Write("&nbsp;")									
							end if
							%>
                          </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then
								response.Write(m_acumul)
							else
								response.Write("&nbsp;")	
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then																				
								response.Write(res1)		
							else
								libera_resultado="n"
								response.Write("&nbsp;")												
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr5="s" and showprova5="s" then												
								response.Write(va_m35_exibe)
							else
								response.Write("&nbsp;")								
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" then				
								response.Write(m2)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" and res1<>"APR" then					
								response.Write(res2)
							else
								response.Write("&nbsp;")									
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr6="s" and showprova6="s" then
								response.Write(m3)
							else
								response.Write("&nbsp;")	
							end if
							%>
                            </div></td>
                        <td width="68" class="<%RESPONSE.Write(WRK_CLASSE)%>"> 
                          <div align="center"><%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" and showapr5="s" and showprova5="s" and showapr6="s" and showprova6="s" then													
								response.Write(res3)	
							else
								response.Write("&nbsp;")									
							end if

							%>
                            </div></td>
							  </tr>						  
					<%
						end if
					end if
					check=check+1
					RSprog.MOVENEXT
					wend

			vetor_resultados= split(vetor_temp_aluno,"#!#")						
			for vr=0 to ubound(vetor_resultados)
				resultado=vetor_resultados(vr)
				
				if resultado="" or isnull(resultado) or resultado="&nbsp;" or resultado=" " or libera_resultado="n" then
					libera_resultado="n"
				else
					if result_temp="REP" then
					else
						if result_temp="REC" then
							if resultado="REP" then	
								result_temp=resultado
							end if			
						else
							if result_temp="PFI" then	
								if resultado="REP" or resultado="REC" then	
									result_temp=resultado
								end if					
							else	
								result_temp=resultado
							end if
						end if	
						if resultado="REC" then
							qtd_rec = qtd_rec+1
						end if						
					end if					
				End if										
			Next
			curso=curso*1
			etapa=etapa*1

			if curso = 1 and etapa<6 then
				if qtd_rec>=3 then
					resultado_aluno="REP"
				else
					resultado_aluno=result_temp			
				end if	
			elseif curso = 1 and etapa>5 then
				if qtd_rec>=4 then
					resultado_aluno="REP"
				else
					resultado_aluno=result_temp			
				end if				
			else
				resultado_aluno=result_temp					
			end if	

				
				Set RSF = Server.CreateObject("ADODB.Recordset")
				SQL_N = "Select * from TB_Frequencia_Periodo WHERE CO_Matricula = "& cod
				Set RSF = CON_N.Execute(SQL_N)
				
				if RSF.eof THEN
				f1="&nbsp;"
				f2="&nbsp;"
				f3="&nbsp;"
				f4="&nbsp;"			
				else	
				f1=RSF("NU_Faltas_P1")
				f2=RSF("NU_Faltas_P2")
				f3=RSF("NU_Faltas_P3")
				f4=RSF("NU_Faltas_P4")		
				END IF		
				%>
						  <tr valign="bottom"> 
							<td height="20" colspan="12"> <div align="right"> 
								
                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr valign="middle"> 
                              <td width="250" height="20" valign="baseline"><font class="style1">Freq&uuml;&ecirc;ncia 
                                (Faltas):</font></td>
									
                              <td width="70" height="20" valign="baseline"> 
                                <div align="right"><font class="style1">Bimestre 
										1:</font></div></td>
									
                              <td width="30" height="20" valign="baseline"><font class="style1"> 
                                <%response.Write(f1)%>
                                </font> </td>
									
                              <td width="70" height="20" valign="baseline"> 
                                <div align="right"><font class="style1">Bimestre 
										2:</font></div></td>
									
                              <td width="30" height="20" valign="baseline"><font class="style1"> 
                                <%response.Write(f2)%>
                                </font> </td>
									
                              <td width="70" height="20" valign="baseline"> 
                                <div align="right"><font class="style1">Bimestre 
										3:</font></div></td>
									
                              <td width="30" height="20" valign="baseline"><font class="style1"> 
                                <%response.Write(f3)%>
                                </font> </td>
									
                              <td width="70" height="20" valign="baseline"> 
                                <div align="right"><font class="style1">Bimestre 
										4:</font></div></td>
									
                              <td width="30" height="20" valign="baseline"><font class="style1"> 
                                <%response.Write(f4)%>
                                </font> </td>
									
                          <td width="450" height="20" valign="baseline">&nbsp; </td>
								  </tr>
								</table>
							  </div></td>
						  </tr> 
                    </table>	-->
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="252" rowspan="2" class="tb_subtit"><div align="left"><strong>Disciplina</strong></div></td>
                      <td width="748" colspan="11" class="tb_subtit"><div align="center"></div>
                        <div align="center">Aproveitamento</div></td>
                    </tr>
                    <tr>
                      <td width="68" class="tb_subtit"><div align="center">BIM 
                          1</div></td>
                      <td width="68" class="tb_subtit"><div align="center">BIM 
                          2</div></td>
                      <td width="68" class="tb_subtit"><div align="center">BIM 
                          3</div></td>
                      <td width="68" class="tb_subtit"><div align="center">BIM 
                          4</div></td>
                      <td width="68" class="tb_subtit"><div align="center">M&eacute;dia 
                          Anual</div></td>
                      <td width="68" class="tb_subtit"><div align="center">Result</div></td>
                      <td width="68" class="tb_subtit"><div align="center">Prova 
                          Final</div></td>
                      <td width="68" class="tb_subtit"><div align="center">M&eacute;dia 
                          Final</div></td>
                      <td width="68" class="tb_subtit"><div align="center">Result</div></td>
                      <td width="68" class="tb_subtit"><div align="center">Prova 
                          Recup</div></td>
                      <td width="68" class="tb_subtit"><div align="center">Result</div></td>
                    </tr>
                    <%
			rec_lancado="sim"
		
			Set RSprog = Server.CreateObject("ADODB.Recordset")
			SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
			RSprog.Open SQLprog, CON0
		
			check=2
				
			while not RSprog.EOF
			
					dividendo1=0
					divisor1=0
					dividendo2=0
					divisor2=0
					dividendo3=0
					divisor3=0
					dividendo4=0
					divisor4=0
					dividendo_ma=0
					divisor_ma=0
					dividendo5=0
					divisor5=0
					dividendo_mf=0
					divisor_mf=0
					dividendo6=0
					divisor6=0
					dividendo_rec=0
					divisor_rec=0
					res1="&nbsp;"
					res2="&nbsp;"
					res3="&nbsp;"
					m2="&nbsp;"
					m3="&nbsp;"										
			
				materia=RSprog("CO_Materia")
				mae=RSprog("IN_MAE")
				fil=RSprog("IN_FIL")
				in_co=RSprog("IN_CO")
				nu_peso=RSprog("NU_Peso")
				ordem=RSprog("NU_Ordem_Boletim")
				
				if mae=TRUE THEN
				
					Set RS1a = Server.CreateObject("ADODB.Recordset")
					SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia_Principal='"&materia&"' order by NU_Ordem_Boletim"
					RS1a.Open SQL1a, CON0
					
				if RS1a.EOF then
				
						Set RS1b = Server.CreateObject("ADODB.Recordset")
						SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"' "
						RS1b.Open SQL1b, CON0
							
						no_materia=RS1b("NO_Materia")
					
						if check mod 2 =0 then
						cor = "tb_fundo_linha_par" 
						cor2 = "tb_fundo_linha_impar" 				
						else 
						cor ="tb_fundo_linha_impar"
						cor2 = "tb_fundo_linha_par" 
						end if					
							
						Set CON_N = Server.CreateObject("ADODB.Connection") 
						ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
						CON_N.Open ABRIRn
					
						for periodofil=1 to 6
					
								
							Set RSnFIL = Server.CreateObject("ADODB.Recordset")
							Set RS3 = Server.CreateObject("ADODB.Recordset")
							SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' AND NU_Periodo="&periodofil
							Set RS3 = CON_N.Execute(SQL_N)
								
							if RS3.EOF then
						if periodofil=1 then
						va_m31="&nbsp;"
						va_m31_exibe="&nbsp;"
						elseif periodofil=2 then
						va_m32="&nbsp;"
						va_m32_exibe="&nbsp;"
						elseif periodofil=3 then
						va_m33="&nbsp;"
						va_m33_exibe="&nbsp;"
						elseif periodofil=4 then
						va_m34="&nbsp;"
						va_m34_exibe="&nbsp;"
						elseif periodofil=5 then
						va_m35="&nbsp;"
						va_m35_exibe="&nbsp;"
						elseif periodofil=6 then
						va_m36="&nbsp;"
						va_m36_exibe="&nbsp;"
						end if	
					else
						if periodofil=1 then
						va_m31=RS3("VA_Media3")
						va_m31_exibe=RS3("VA_Media3")
						elseif periodofil=2 then
						va_m32=RS3("VA_Media3")
						va_m32_exibe=RS3("VA_Media3")
						elseif periodofil=3 then
						va_m33=RS3("VA_Media3")
						va_m33_exibe=RS3("VA_Media3")
						elseif periodofil=4 then
						va_m34=RS3("VA_Media3")
						va_m34_exibe=RS3("VA_Media3")
						elseif periodofil=5 then
						va_m35=RS3("VA_Media3")
						va_m35_exibe=RS3("VA_Media3")
						elseif periodofil=6 then
						va_m36=RS3("VA_Media3")
						va_m36_exibe=RS3("VA_Media3")
						end if
					end if
							
							if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
							dividendo1=0
							divisor1=0
							else
							dividendo1=va_m31
							divisor1=1
							if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
																if va_m31 > 90 then
							va_m31_exibe="E"
							elseif (va_m31 > 70) and (va_m31 <= 90) then
							va_m31_exibe="MB"
							elseif (va_m31 > 60) and (va_m31 <= 70) then							
							va_m31_exibe="B"
							elseif (va_m31 > 49) and (va_m31 <= 60) then
							va_m31_exibe="R"
							else							
							va_m31_exibe="I"
							end if													
								end if							
							end if	
							
							if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
							dividendo2=0
							divisor2=0
							else
												dividendo2=va_m32
					divisor2=1
																						if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then	
														if va_m32 > 90 then
									va_m32_exibe="E"
									elseif (va_m32 > 70) and (va_m32 <= 90) then
									va_m32_exibe="MB"
									elseif (va_m32 > 60) and (va_m32 <= 70) then							
									va_m32_exibe="B"
									elseif (va_m32 > 49) and (va_m32 <= 60) then
									va_m32_exibe="R"
									else							
									va_m32_exibe="I"
								end if
						end if
							end if
							
							if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
							dividendo3=0
							divisor3=0
							else
										dividendo3=va_m33
					divisor3=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
								if va_m33 > 90 then
									va_m33_exibe="E"
									elseif (va_m33 > 70) and (va_m33 <= 90) then
									va_m33_exibe="MB"
									elseif (va_m33 > 60) and (va_m33 <= 70) then							
									va_m33_exibe="B"
									elseif (va_m33 > 49) and (va_m33 <= 60) then
									va_m33_exibe="R"
									else							
									va_m33_exibe="I"
								end if
						end if
							end if
							
							if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
							dividendo4=0
							divisor4=0
							else
														dividendo4=va_m34
							divisor4=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then					
								if va_m34 > 90 then
									va_m34_exibe="E"
									elseif (va_m34 > 70) and (va_m34 <= 90) then
									va_m34_exibe="MB"
									elseif (va_m34 > 60) and (va_m34 <= 70) then							
									va_m34_exibe="B"
									elseif (va_m34 > 49) and (va_m34 <= 60) then
									va_m34_exibe="R"
									else							
									va_m34_exibe="I"
								end if
						end if
							end if
										
							dividendo_ma=dividendo1+dividendo2+dividendo3+dividendo4
							divisor_ma=divisor1+divisor2+divisor3+divisor4
							
											
							if divisor_ma<4 then
							ma="&nbsp;"
							else
							ma=dividendo_ma/divisor_ma
							end if
								
							if ma="&nbsp;" then
							else
							'mf=mf/10
								decimo = ma - Int(ma)
									If decimo >= 0.5 Then
										nota_arredondada = Int(ma) + 1
										ma=nota_arredondada
									Else
										nota_arredondada = Int(ma)
										ma=nota_arredondada					
									End If
								ma = formatNumber(ma,0)
								ma=ma*1								
'								if ma>67 and ma<70then
'									ma=70
'								end if								
		
							end if
										
										
							ma = AcrescentaBonusMediaAnual(cod, materia, ma)
																
							if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
							nota_aux_m2_1="&nbsp;"
		
							else
							nota_aux_m2_1=va_m35
							end if
		
						
							if isnull(va_m36) or va_m36="&nbsp;" or va_m36="" then
							nota_aux_m3_1="&nbsp;"
							else
							nota_aux_m3_1=va_m36
							end if
		
									showapr1="s"
		
									showprova1="s"
		
									showapr2="s"
		
									showprova2="s"
		
									showapr3="s"
		
									showprova3="s"
		
									showapr4="s"
		
									showprova4="s"
		
						NEXT
						
							if ma="&nbsp;" then
							else	
												
								resultados=novo_regra_aprovacao (cod, materia,curso,etapa,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,"waboletim")
								medias_resultados=split(resultados,"#!#")

								res1=medias_resultados(1)
								res2=medias_resultados(3)
								res3=medias_resultados(5)
								m2=medias_resultados(2)
								m3=medias_resultados(4)
								
								'Se a coluna for de resultado e o resultado estiver preenchido
								'Verifica se o aluno foi aprovado pelo conselho de classe
								if res1<>"&nbsp;" then
									tipo_media = "MA"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)							
									if modifica_result <> "N" then
										res1 = modifica_result
									end if		
								end if	
									
								if res2<>"&nbsp;" then
									tipo_media = "RF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res2 = modifica_result
									end if																										
								end if															
								if res3<>"&nbsp;" then
									tipo_media = "MF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res3 = modifica_result
									end if	
								end if										
					
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
									if ma<>"&nbsp;" then
										if ma > 90 then
										ma="E"
										elseif (ma > 70) and (ma <= 90) then
										ma="MB"
										elseif (ma > 60) and (ma <= 70) then							
										ma="B"
										elseif (ma > 49) and (ma <= 60) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 > 90 then
										m2="E"
										elseif (m2 > 70) and (m2 <= 90) then
										m2="MB"
										elseif (m2 > 60) and (m2 <= 70) then							
										m2="B"
										elseif (m2 > 49) and (m2 <= 60) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if isnull(m3) or m3="" then
									'	m3=m2
									else
										if m3<>"&nbsp;" then									
											if m3 > 90 then
											m3="E"
											elseif (m3 > 70) and (m3 <= 90) then
											m3="MB"
											elseif (m3 > 60) and (m3 <= 70) then							
											m3="B"
											elseif (m3 > 49) and (m3 <= 60) then
											m3="R"
											else							
											m3="I"
											end if	
										end if						
									end if				
					
								end if							
							
							end if
					%>
                    <tr>
                      <td width="252" class="<%response.Write(cor)%>"><%response.Write(no_materia)
								  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
								  %></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then																	
									response.Write(va_m31_exibe)
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then												
									response.Write(va_m32_exibe)						
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then					
									response.Write(va_m33_exibe)
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then					
									response.Write(va_m34_exibe)
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then
									response.Write(ma)
									else
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then
																						
									response.Write(res1)					
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then												
									response.Write(va_m35_exibe)
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then					
									response.Write(m2)
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then					
									response.Write(res2)
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then
									response.Write(m3)
									else
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr3="s" and showprova3="s" then													
									response.Write(res3)	
									end if
		
									%>
                        </div></td>
                    </tr>
                    <%
				else
						Set RS1b = Server.CreateObject("ADODB.Recordset")
						SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"' order by NU_Ordem_Boletim"
						RS1b.Open SQL1b, CON0
							
						no_materia=RS1b("NO_Materia")
					
						if check mod 2 =0 then
						cor = "tb_fundo_linha_par" 
						cor2 = "tb_fundo_linha_impar" 				
						else 
						cor ="tb_fundo_linha_impar"
						cor2 = "tb_fundo_linha_par" 
						end if					
							
						Set CON_N = Server.CreateObject("ADODB.Connection") 
						ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
						CON_N.Open ABRIRn
					
						for periodofil=1 to 6
					
								
							Set RSnFIL = Server.CreateObject("ADODB.Recordset")
							Set RS3 = Server.CreateObject("ADODB.Recordset")
							SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' AND NU_Periodo="&periodofil
							Set RS3 = CON_N.Execute(SQL_N)
								
							if RS3.EOF then
						if periodofil=1 then
						va_m31="&nbsp;"
						va_m31_exibe="&nbsp;"
						elseif periodofil=2 then
						va_m32="&nbsp;"
						va_m32_exibe="&nbsp;"
						elseif periodofil=3 then
						va_m33="&nbsp;"
						va_m33_exibe="&nbsp;"
						elseif periodofil=4 then
						va_m34="&nbsp;"
						va_m34_exibe="&nbsp;"
						elseif periodofil=5 then
						va_m35="&nbsp;"
						va_m35_exibe="&nbsp;"
						elseif periodofil=6 then
						va_m36="&nbsp;"
						va_m36_exibe="&nbsp;"
						end if	
					else
						if periodofil=1 then
						va_m31=RS3("VA_Media3")
						va_m31_exibe=RS3("VA_Media3")
						elseif periodofil=2 then
						va_m32=RS3("VA_Media3")
						va_m32_exibe=RS3("VA_Media3")
						elseif periodofil=3 then
						va_m33=RS3("VA_Media3")
						va_m33_exibe=RS3("VA_Media3")
						elseif periodofil=4 then
						va_m34=RS3("VA_Media3")
						va_m34_exibe=RS3("VA_Media3")
						elseif periodofil=5 then
						va_m35=RS3("VA_Media3")
						va_m35_exibe=RS3("VA_Media3")
						elseif periodofil=6 then
						va_m36=RS3("VA_Media3")
						va_m36_exibe=RS3("VA_Media3")
						end if
					end if
							
							if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
							dividendo1=0
							divisor1=0
							else
							dividendo1=va_m31
							divisor1=1
							if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
								if va_m31 > 90 then
									va_m31_exibe="E"
								elseif (va_m31 > 70) and (va_m31 <= 90) then
									va_m31_exibe="MB"
								elseif (va_m31 > 60) and (va_m31 <= 70) then							
									va_m31_exibe="B"
								elseif (va_m31 > 49) and (va_m31 <= 60) then
									va_m31_exibe="R"
								else							
								va_m31_exibe="I"
								end if													
						end if							
							end if	
							
							if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
							dividendo2=0
							divisor2=0
							else
												dividendo2=va_m32
					divisor2=1
						if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then	
														if va_m32 > 90 then
									va_m32_exibe="E"
									elseif (va_m32 > 70) and (va_m32 <= 90) then
									va_m32_exibe="MB"
									elseif (va_m32 > 60) and (va_m32 <= 70) then							
									va_m32_exibe="B"
									elseif (va_m32 > 49) and (va_m32 <= 60) then
									va_m32_exibe="R"
									else							
									va_m32_exibe="I"
								end if
						end if
							end if
							
							if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
							dividendo3=0
							divisor3=0
							else
										dividendo3=va_m33
					divisor3=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
								if va_m33 > 90 then
									va_m33_exibe="E"
									elseif (va_m33 > 70) and (va_m33 <= 90) then
									va_m33_exibe="MB"
									elseif (va_m33 > 60) and (va_m33 <= 70) then							
									va_m33_exibe="B"
									elseif (va_m33 > 49) and (va_m33 <= 60) then
									va_m33_exibe="R"
									else							
									va_m33_exibe="I"
								end if
						end if
							end if
							
							if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
							dividendo4=0
							divisor4=0
							else
														dividendo4=va_m34
							divisor4=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then					
								if va_m34 > 90 then
									va_m34_exibe="E"
									elseif (va_m34 > 70) and (va_m34 <= 90) then
									va_m34_exibe="MB"
									elseif (va_m34 > 60) and (va_m34 <= 70) then							
									va_m34_exibe="B"
									elseif (va_m34 > 49) and (va_m34 <= 60) then
									va_m34_exibe="R"
									else							
									va_m34_exibe="I"
								end if
						end if
							end if
										
							dividendo_ma=dividendo1+dividendo2+dividendo3+dividendo4
							divisor_ma=divisor1+divisor2+divisor3+divisor4
											
							if divisor_ma<4 then
							ma="&nbsp;"
							else
							ma=dividendo_ma/divisor_ma
							end if
							
							if ma="&nbsp;" then
							else
							'mf=mf/10
								decimo = ma - Int(ma)
									If decimo >= 0.5 Then
										nota_arredondada = Int(ma) + 1
										ma=nota_arredondada
									Else
										nota_arredondada = Int(ma)
										ma=nota_arredondada					
									End If
								ma = formatNumber(ma,0)
								ma=ma*1								
'								if ma>67 and ma<70then
'									ma=70
'								end if
							end if
											
											
							ma = AcrescentaBonusMediaAnual(cod, materia, ma)
																
							if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
							nota_aux_m2_1="&nbsp;"
		
							else
							nota_aux_m2_1=va_m35
							end if
		
						
							if isnull(va_m36) or va_m36="&nbsp;" or va_m36="" then
							nota_aux_m3_1="&nbsp;"
							else
							nota_aux_m3_1=va_m36
							end if
		
									showapr1="s"
		
									showprova1="s"
		
									showapr2="s"
		
									showprova2="s"
		
									showapr3="s"
		
									showprova3="s"
		
									showapr4="s"
		
									showprova4="s"
		
						NEXT
						
							if ma="&nbsp;" then
							else	
												
								resultados=novo_regra_aprovacao (cod, materia,curso,etapa,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,"waboletim")
								medias_resultados=split(resultados,"#!#")
								
								res1=medias_resultados(1)
								res2=medias_resultados(3)
								res3=medias_resultados(5)
								m2=medias_resultados(2)
								m3=medias_resultados(4)
								
								'Se a coluna for de resultado e o resultado estiver preenchido
								'Verifica se o aluno foi aprovado pelo conselho de classe								
								if res1<>"&nbsp;" then
									tipo_media = "MA"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res1 = modifica_result
									end if		
								end if	
								if res2<>"&nbsp;" then
									tipo_media = "RF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res2 = modifica_result
									end if																										
								end if															
								if res3<>"&nbsp;" then
									tipo_media = "MF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res3 = modifica_result
									end if	
								end if										
						
							
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then									
									if ma<>"&nbsp;" then
										if ma > 90 then
										ma="E"
										elseif (ma > 70) and (ma <= 90) then
										ma="MB"
										elseif (ma > 60) and (ma <= 70) then							
										ma="B"
										elseif (ma > 49) and (ma <= 60) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 > 90 then
										m2="E"
										elseif (m2 > 70) and (m2 <= 90) then
										m2="MB"
										elseif (m2 > 60) and (m2 <= 70) then							
										m2="B"
										elseif (m2 > 49) and (m2 <= 60) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if isnull(m3) or m3="" then
									'	m3=m2
									else
										if m3<>"&nbsp;" then									
											if m3 > 90 then
											m3="E"
											elseif (m3 > 70) and (m3 <= 90) then
											m3="MB"
											elseif (m3 > 60) and (m3 <= 70) then							
											m3="B"
											elseif (m3 > 49) and (m3 <= 60) then
											m3="R"
											else							
											m3="I"
											end if	
										end if						
									end if						
											
								end if						
							end if
					%>
                    <tr>

                      <td width="252" class="<%response.Write(cor)%>"><%response.Write(no_materia)
								  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
								  %></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%									if showapr1="s" then																	
									response.Write(va_m31_exibe)
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then												
									response.Write(va_m32_exibe)						
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then					
									response.Write(va_m33_exibe)
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then					
									response.Write(va_m34_exibe)
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then
									response.Write(ma)
									else
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then
																						
									response.Write(res1)					
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then												
									response.Write(va_m35_exibe)
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then					
									response.Write(m2)
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then					
									response.Write(res2)
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then
									response.Write(m3)
									else
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr3="s" and showprova3="s" then													
									response.Write(res3)	
									end if
		
									%>
                        </div></td>
                    </tr>
                    <%					
							
								divisor_m_acumul=0
								peso_acumula=0
								acumula_m1=0
								m31_ac=0
								m32_ac=0			
								m33_ac=0
								m34_ac=0
								m35_ac=0
								m36_ac=0
								m31_exibe=0
								m32_exibe=0
								m33_exibe=0
								m34_exibe=0
								m35_exibe=0
								m36_exibe=0								
								nu_peso_fil=0
								dividendo1=0
								dividendo2=0
								dividendo3=0
								dividendo4=0
								dividendo5=0
								dividendo6=0
								conta_fil=0
								conta1=0
								conta2=0
								conta3=0
								conta4=0
								conta5=0
								conta6=0
								while not RS1a.EOF
								conta_fil=conta_fil+1
							
									materia_fil=RS1a("CO_Materia")
								
											Set RS1b = Server.CreateObject("ADODB.Recordset")
											SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia_fil&"' order by NU_Ordem_Boletim"
											RS1b.Open SQL1b, CON0
											
									no_materia_fil=RS1b("NO_Materia")
									
									Set RSpa = Server.CreateObject("ADODB.Recordset")
									SQLpa= "SELECT * FROM TB_Programa_Aula where CO_Materia='"&materia_fil&"' AND CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"'"
									RSpa.Open SQLpa, CON0
															
									nu_peso_fil=RSpa("NU_Peso")	
									
									if isnull(nu_peso_fil) or nu_peso_fil="" then
										nu_peso_fil=1
									end if					
							
							for periodofil=1 to 6	
										
											Set RSnFIL = Server.CreateObject("ADODB.Recordset")
											Set RS3 = Server.CreateObject("ADODB.Recordset")
											SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia_Principal = '"& materia &"' AND CO_Materia = '"& materia_fil &"' and NU_Periodo="&periodofil
											Set RS3 = CON_N.Execute(SQL_N)
										  						  								
							if RS3.EOF then
						if periodofil=1 then
						va_m31="&nbsp;"
						va_m31_exibe="&nbsp;"
						conta1=conta1
						elseif periodofil=2 then
						va_m32="&nbsp;"
						va_m32_exibe="&nbsp;"
						conta2=conta2
						elseif periodofil=3 then
						va_m33="&nbsp;"
						va_m33_exibe="&nbsp;"
						conta3=conta3
						elseif periodofil=4 then
						va_m34="&nbsp;"
						va_m34_exibe="&nbsp;"
						conta4=conta4
						elseif periodofil=5 then
						va_m35="&nbsp;"
						va_m35_exibe="&nbsp;"
						conta5=conta5
						elseif periodofil=6 then
						va_m36="&nbsp;"
						va_m36_exibe="&nbsp;"
						conta6=conta6
						end if	
					else
						if periodofil=1 then
						va_m31=RS3("VA_Media3")
						va_m31_exibe=RS3("VA_Media3")
								if isnull(va_m31_exibe) or va_m31_exibe="" then
								conta1=conta1
								else
								conta1=conta1+1
								end if								
						elseif periodofil=2 then
						va_m32=RS3("VA_Media3")
						va_m32_exibe=RS3("VA_Media3")
								if isnull(va_m32_exibe) or va_m32_exibe="" then
								conta2=conta2
								else
								conta2=conta2+1
								end if						
						elseif periodofil=3 then
						va_m33=RS3("VA_Media3")
						va_m33_exibe=RS3("VA_Media3")
								if isnull(va_m33_exibe) or va_m33_exibe="" then
								conta3=conta3
								else
								conta3=conta3+1
								end if
						elseif periodofil=4 then
						va_m34=RS3("VA_Media3")
						va_m34_exibe=RS3("VA_Media3")
								if isnull(va_m34_exibe) or va_m34_exibe="" then
								conta4=conta4
								else
								conta4=conta4+1
								end if						
						elseif periodofil=5 then
						va_m35=RS3("VA_Media3")
						va_m35_exibe=RS3("VA_Media3")
								if isnull(va_m35_exibe) or va_m35_exibe="" then
								conta5=conta5
								else
								conta5=conta5+1
								end if						
						elseif periodofil=6 then
						va_m36=RS3("VA_Media3")
						va_m36_exibe=RS3("VA_Media3")
								if isnull(va_m36_exibe) or va_m36_exibe="" then
								conta6=conta6
								else
								conta6=conta6+1
								end if						
						end if
					end if

						if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
							dividendo1=0
							divisor1=0
						else
							dividendo1=va_m31
							divisor1=1
							if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
																	if va_m31 > 90 then
								va_m31_exibe="E"
								elseif (va_m31 > 70) and (va_m31 <= 90) then
								va_m31_exibe="MB"
								elseif (va_m31 > 60) and (va_m31 <= 70) then							
								va_m31_exibe="B"
								elseif (va_m31 > 49) and (va_m31 <= 60) then
								va_m31_exibe="R"
								else							
								va_m31_exibe="I"
								end if													
							end if							
						end if	
							
						if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
							dividendo2=0
							divisor2=0
						else
							dividendo2=va_m32
							divisor2=1
							if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then	
							
								if va_m32 > 90 then
									va_m32_exibe="E"
								elseif (va_m32 > 70) and (va_m32 <= 90) then
									va_m32_exibe="MB"
								elseif (va_m32 > 60) and (va_m32 <= 70) then							
									va_m32_exibe="B"
								elseif (va_m32 > 49) and (va_m32 <= 60) then
									va_m32_exibe="R"
								else							
									va_m32_exibe="I"
							end if
						end if
							end if
							
							if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
							dividendo3=0
							divisor3=0
							else
							dividendo3=va_m33
							divisor3=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then
									if va_m33 > 90 then
									va_m33_exibe="E"
									elseif (va_m33 > 70) and (va_m33 <= 90) then
									va_m33_exibe="MB"
									elseif (va_m33 > 60) and (va_m33 <= 70) then							
									va_m33_exibe="B"
									elseif (va_m33 > 49) and (va_m33 <= 60) then
									va_m33_exibe="R"
									else							
									va_m33_exibe="I"
								end if
								end if
							end if
							
							if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
							dividendo4=0
							divisor4=0
							else
							dividendo4=va_m34
							divisor4=1
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then					
									if va_m34 > 90 then
									va_m34_exibe="E"
									elseif (va_m34 > 70) and (va_m34 <= 90) then
									va_m34_exibe="MB"
									elseif (va_m34 > 60) and (va_m34 <= 70) then							
									va_m34_exibe="B"
									elseif (va_m34 > 49) and (va_m34 <= 60) then
									va_m34_exibe="R"
									else							
									va_m34_exibe="I"
								end if
								end if
							end if
										
							dividendo_ma=dividendo1+dividendo2+dividendo3+dividendo4
							divisor_ma=divisor1+divisor2+divisor3+divisor4
											
							if divisor_ma<4 then
							ma="&nbsp;"
							else
							ma=dividendo_ma/divisor_ma
							end if
							
							if ma="&nbsp;" then
							else
							'mf=mf/10
								decimo = ma - Int(ma)
									If decimo >= 0.5 Then
										nota_arredondada = Int(ma) + 1
										ma=nota_arredondada
									Else
										nota_arredondada = Int(ma)
										ma=nota_arredondada					
									End If
								ma = formatNumber(ma,0)
								ma=ma*1
'								if ma>67 and ma<70then
'									ma=70
'								end if		
							end if
							
							ma = AcrescentaBonusMediaAnual(cod, materia, ma)																
							
							if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
							nota_aux_m2_1="&nbsp;"
							dividendo5=0
							else
							nota_aux_m2_1=va_m35
							dividendo5=va_m35
							end if
		
						
							if isnull(va_m36) or va_m36="&nbsp;" or va_m36="" then
							nota_aux_m3_1="&nbsp;"
							dividendo6=0
							else
							nota_aux_m3_1=va_m36
							dividendo6=va_m36
							end if
		
									showapr1="s"
		
									showprova1="s"
		
									showapr2="s"
		
									showprova2="s"
		
									showapr3="s"
		
									showprova3="s"
		
									showapr4="s"
		
									showprova4="s"
		
						NEXT
					
							if ma="&nbsp;" then
							else	
'response.Write(materia&":"&ma&","&nota_aux_m2_1&","&nota_aux_m3_1&"<BR>")													
'								resultados=novo_regra_aprovacao (cod, materia,curso,etapa,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,"waboletim")
''response.Write(materia&":"&resultados&"<BR>")									
'								medias_resultados=split(resultados,"#!#")
'								
'								res1=medias_resultados(1)
'								res2=medias_resultados(3)
'								res3=medias_resultados(5)
'								m2=medias_resultados(2)
'								m3=medias_resultados(4)	
								
								'Se a coluna for de resultado e o resultado estiver preenchido
								'Verifica se o aluno foi aprovado pelo conselho de classe
								if res1<>"&nbsp;" then
									tipo_media = "MA"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res1 = modifica_result
									end if		
								end if	
								if res2<>"&nbsp;" then
									tipo_media = "RF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res2 = modifica_result
									end if																										
								end if															
								if res3<>"&nbsp;" then
									tipo_media = "MF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res3 = modifica_result
									end if	
								end if										
							
														
						
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then									
									if ma<>"&nbsp;" then
										if ma > 90 then
										ma="E"
										elseif (ma > 70) and (ma <= 90) then
										ma="MB"
										elseif (ma > 60) and (ma <= 70) then							
										ma="B"
										elseif (ma > 49) and (ma <= 60) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 > 90 then
										m2="E"
										elseif (m2 > 70) and (m2 <= 90) then
										m2="MB"
										elseif (m2 > 60) and (m2 <= 70) then							
										m2="B"
										elseif (m2 > 49) and (m2 <= 60) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if isnull(m3) or m3="" then
									'	m3=m2
									else
										if m3<>"&nbsp;" then									
											if m3 > 90 then
											m3="E"
											elseif (m3 > 70) and (m3 <= 90) then
											m3="MB"
											elseif (m3 > 60) and (m3 <= 70) then							
											m3="B"
											elseif (m3 > 49) and (m3 <= 60) then
											m3="R"
											else							
											m3="I"
											end if	
										end if						
									end if					
											
								end if							
							
							
							end if
					%>
                    <tr>
                      <td width="252" class="<%response.Write(cor)%>">&nbsp;&nbsp;&nbsp;
                        <%response.Write(no_materia_fil)
								  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
								  %></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%									if showapr1="s" then																	
									response.Write(va_m31_exibe)
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then												
									response.Write(va_m32_exibe)						
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then					
									response.Write(va_m33_exibe)
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then					
									response.Write(va_m34_exibe)
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then
									response.Write(ma)
									else
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then
																						
									'response.Write(res1)					
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then												
									response.Write(va_m35_exibe)
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then					
									response.Write(va_m35)
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then					
									'response.Write(res2)
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%							
									if showapr2="s" and showprova2="s" then
									response.Write(va_m36)
									else
									end if
									%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr3="s" and showprova3="s" then													
									'response.Write(res3)	
									end if
		
									%>
                        </div></td>
                    </tr>
                    <%		if in_co=TRUE then
								peso_acumula=nu_peso_fil					
							else															
								peso_acumula=peso_acumula+nu_peso_fil
							end if	
							m31_ac=m31_ac+(dividendo1*nu_peso_fil)	
							m32_ac=m32_ac+(dividendo2*nu_peso_fil)
							m33_ac=m33_ac+(dividendo3*nu_peso_fil)
							m34_ac=m34_ac+(dividendo4*nu_peso_fil)							
							m35_ac=m35_ac+(dividendo5*nu_peso_fil)
							m36_ac=m36_ac+(dividendo6*nu_peso_fil)
							RS1a.movenext
							wend
							
							conta1=conta1*1
							conta2=conta2*1
							conta3=conta3*1
							conta4=conta4*1
							conta5=conta5*1
							conta6=conta6*1																																			
							if conta1<conta_fil then
								m31_exibe="&nbsp;"							
							else
								m31_exibe=m31_ac/peso_acumula								
								decimo = m31_exibe - Int(m31_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m31_exibe) + 1
									m31_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m31_exibe)
									m31_exibe=nota_arredondada					
								End If	
								if m31_exibe>100 then
									m31_exibe=100
								end if															
								m31_exibe = formatNumber(m31_exibe,0)		
							end if
													
							if conta2<conta_fil then
								m32_exibe="&nbsp;"							
							else
								m32_exibe=m32_ac/peso_acumula
								decimo = m32_exibe - Int(m32_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m32_exibe) + 1
									m32_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m32_exibe)
									m32_exibe=nota_arredondada					
								End If		
								if m32_exibe>100 then
									m32_exibe=100
								end if													
								m32_exibe = formatNumber(m32_exibe,0)						
							end if							
							
							if conta3<conta_fil then
								m33_exibe="&nbsp;"							
							else
								m33_exibe=m33_ac/peso_acumula
								decimo = m33_exibe - Int(m33_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m33_exibe) + 1
									m33_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m33_exibe)
									m33_exibe=nota_arredondada					
								End If	
								if m33_exibe>100 then
									m33_exibe=100
								end if														
								m33_exibe = formatNumber(m33_exibe,0)						
							end if
							
							if conta4<conta_fil then
								m34_exibe="&nbsp;"							
							else
								m34_exibe=m34_ac/peso_acumula
								decimo = m34_exibe - Int(m34_exibe)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m34_exibe) + 1
									m34_exibe=nota_arredondada
								Else
									nota_arredondada = Int(m34_exibe)
									m34_exibe=nota_arredondada					
								End If				
								if m34_exibe>100 then
									m34_exibe=100
								end if											
								m34_exibe = formatNumber(m34_exibe,0)				
							end if
							
							'response.Write(conta5&"<"&conta_fil)
							if conta5<conta_fil then
								m35_mae="&nbsp;"							
							else
								m35_mae=m35_ac/peso_acumula
								decimo = m35_mae - Int(m35_mae)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m35_mae) + 1
									m35_mae=nota_arredondada
								Else
									nota_arredondada = Int(m35_mae)
									m35_mae=nota_arredondada					
								End If	
								if m35_mae>100 then
									m35_mae=100
								end if														
								m35_mae = formatNumber(m35_mae,0)			
							end if																					
							
							if conta6<conta_fil then
								m36_mae="&nbsp;"							
							else
							
								m36_mae=m36_ac/peso_acumula
								decimo = m36_mae - Int(m36_mae)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m36_mae) + 1
									m36_mae=nota_arredondada
								Else
									nota_arredondada = Int(m36_mae)
									m36_mae=nota_arredondada					
								End If				
								if m36_mae>100 then
									m36_mae=100
								end if											
								m36_mae = formatNumber(m36_mae,0)				
							end if
							
							m31_mae=m31_exibe																																				
							m32_mae=m32_exibe	
							m33_mae=m33_exibe								
							m34_mae=m34_exibe							
							
							if m31_mae="&nbsp;" then
								divisor_m_acumul=divisor_m_acumul
							else
								divisor_m_acumul=divisor_m_acumul+1
							end if
	
							if m32_mae="&nbsp;" then
								divisor_m_acumul=divisor_m_acumul
							else
								divisor_m_acumul=divisor_m_acumul+1								
							end if
							
							if m33_mae="&nbsp;" then
								divisor_m_acumul=divisor_m_acumul
							else
								divisor_m_acumul=divisor_m_acumul+1								
							end if
							
							if m34_mae="&nbsp;" then
								divisor_m_acumul=divisor_m_acumul
							else
								divisor_m_acumul=divisor_m_acumul+1								
							end if	
										
										
							if isnull(m35_mae) or m35_mae= "" then
								nota_aux_m2_1="&nbsp;"
							else							
								nota_aux_m2_1=m35_mae
							end if	
								
							if isnull(m36_mae) or m36_mae= "" then
								nota_aux_m3_1="&nbsp;"
							else							
								nota_aux_m3_1=m36_mae
							end if								

							
'response.write(dividendo_m_acumul&"-"&m31_ac&"-"&m32_ac&"-"&m33_ac&"-"&m34_ac&"-"&divisor_m_acumul&"-"&nota_aux_m2_1&"-"&nota_aux_m3_1)										
										
							minimo_exibir=4
							'response.Write(va_m31&" - "&va_m32&" - "&va_m33&" - "&va_m34&" - "&divisor_m_acumul&"<"&minimo_exibir)								
							if divisor_m_acumul<minimo_exibir then
								m_acumul="&nbsp;"
							else
								m31_mae=m31_mae*1
								m32_mae=m32_mae*1
								m33_mae=m33_mae*1
								m34_mae=m34_mae*1
								dividendo_m_acumul=m31_mae+m32_mae+m33_mae+m34_mae
								
								m_acumul=dividendo_m_acumul/divisor_m_acumul
							end if
							
							if m_acumul="&nbsp;" then
							else
							'mf=mf/10

								decimo = m_acumul - Int(m_acumul)
'response.write(m_acumul&"-"&decimo)
								'decimo =formatNumber(decimo,1)
'response.write(m_acumul&"-"&decimo)				
									If decimo >= 0.5 Then
										nota_arredondada = Int(m_acumul) + 1
										m_acumul=nota_arredondada
									Else
										nota_arredondada = Int(m_acumul)
										m_acumul=nota_arredondada					
									End If
								m_acumul = formatNumber(m_acumul,0)
								m_acumul =m_acumul *1
'								if m_acumul >67 and m_acumul <70 then
'									m_acumul =70
'								end if	
								if m_acumul>100 then
									m_acumul=100
								end if	
								m_acumul = AcrescentaBonusMediaAnual(cod, materia, m_acumul)
		
							end if							
							
							if m_acumul="&nbsp;" then
							else	

								resultados=novo_regra_aprovacao (cod, materia,curso,etapa,m_acumul,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,"waboletim")
								medias_resultados=split(resultados,"#!#")
								
								res1=medias_resultados(1)
								res2=medias_resultados(3)
								res3=medias_resultados(5)
								m2=medias_resultados(2)
								m3=medias_resultados(4)
								
								'Se a coluna for de resultado e o resultado estiver preenchido
								'Verifica se o aluno foi aprovado pelo conselho de classe
							
								if res1<>"&nbsp;" then
									tipo_media = "MA"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res1 = modifica_result
									end if		
								end if	
								if res2<>"&nbsp;" then
									tipo_media = "RF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res2 = modifica_result
									end if																										
								end if															
								if res3<>"&nbsp;" then
									tipo_media = "MF"
									modifica_result = Verifica_Conselho_Classe(cod, materia, tipo_media, outro)
									if modifica_result <> "N" then
										res3 = modifica_result
									end if	
								end if										
							
								
												
								if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then									
									if ma<>"&nbsp;" then
										if ma > 90 then
										ma="E"
										elseif (ma > 70) and (ma <= 90) then
										ma="MB"
										elseif (ma > 60) and (ma <= 70) then							
										ma="B"
										elseif (ma > 49) and (ma <= 60) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 > 90 then
										m2="E"
										elseif (m2 > 70) and (m2 <= 90) then
										m2="MB"
										elseif (m2 > 60) and (m2 <= 70) then							
										m2="B"
										elseif (m2 > 49) and (m2 <= 60) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if isnull(m3) or m3="" then
									'	m3=m2
									else
										if m3<>"&nbsp;" then									
											if m3 > 90 then
											m3="E"
											elseif (m3 > 70) and (m3 <= 90) then
											m3="MB"
											elseif (m3 > 60) and (m3 <= 70) then							
											m3="B"
											elseif (m3 > 49) and (m3 <= 60) then
											m3="R"
											else							
											m3="I"
											end if	
										end if						
									end if				
											
								end if							
							
							end if							
							%>
                    <tr class="tb_fundo_linha_media">
                      <td width="252" >&nbsp;&nbsp;&nbsp; M&eacute;dia </td>
                      <td width="68" ><div align="center">
                          <%
									if showapr1="s" then																	
									response.Write(m31_exibe)
									end if
									%>
                        </div></td>
                      <td width="68" ><div align="center">
                          <%
									if showapr1="s" then												
									response.Write(m32_exibe)						
									end if
									%>
                        </div></td>
                      <td width="68" ><div align="center">
                          <%
									if showapr1="s" then					
									response.Write(m33_exibe)
									end if
									%>
                        </div></td>
                      <td width="68" ><div align="center">
                          <%
									if showapr1="s" then					
									response.Write(m34_exibe)
									end if
									%>
                        </div></td>
                      <td width="68" ><div align="center">
                          <%
									if showapr1="s" then
									response.Write(m_acumul)
									else
									end if
									%>
                        </div></td>
                      <td width="68" ><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then
																						
									response.Write(res1)					
									end if
									%>
                        </div></td>
                      <td width="68" ><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then												
									response.Write(m35_mae)
									end if
									%>
                        </div></td>
                      <td width="68" ><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then					
									response.Write(m2)
									end if
									%>
                        </div></td>
                      <td width="68" ><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then					
									response.Write(res2)
									end if
									%>
                        </div></td>
                      <td width="68" ><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then
									response.Write(m3)
									else
									end if
									%>
                        </div></td>
                      <td width="68" ><div align="center">
                          <%
									if showapr3="s" and showprova3="s" then													
									response.Write(res3)	
									end if
		
									%>
                        </div></td>
                    </tr>
                    <%
						end if
					end if
					check=check+1
					RSprog.MOVENEXT
					wend


				
				Set RSF = Server.CreateObject("ADODB.Recordset")
				SQL_N = "Select * from TB_Frequencia_Periodo WHERE CO_Matricula = "& cod
				Set RSF = CON_N.Execute(SQL_N)
				
				if RSF.eof THEN
				f1="&nbsp;"
				f2="&nbsp;"
				f3="&nbsp;"
				f4="&nbsp;"			
				else	
				f1=RSF("NU_Faltas_P1")
				f2=RSF("NU_Faltas_P2")
				f3=RSF("NU_Faltas_P3")
				f4=RSF("NU_Faltas_P4")		
				END IF		
				%>
                    <tr valign="bottom">
                      <td height="20" colspan="12"><div align="right">
                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr valign="middle">
                              <td width="250" height="20"><font class="form_dado_texto">Freq&uuml;&ecirc;ncia 
                                (Faltas):</font></td>
                              <td width="70" height="20"><div align="right"><font class="form_dado_texto">Bimestre 
                                  1:</font></div></td>
                              <td width="30" height="20"><font class="form_corpo">
                                <%response.Write(f1)%>
                                </font></td>
                              <td width="70" height="20"><div align="right"><font class="form_dado_texto">Bimestre 
                                  2:</font></div></td>
                              <td width="30" height="20"><font class="form_corpo">
                                <%response.Write(f2)%>
                                </font></td>
                              <td width="70" height="20"><div align="right"><font class="form_dado_texto">Bimestre 
                                  3:</font></div></td>
                              <td width="30" height="20"><font class="form_corpo">
                                <%response.Write(f3)%>
                                </font></td>
                              <td width="70" height="20"><div align="right"><font class="form_dado_texto">Bimestre 
                                  4:</font></div></td>
                              <td width="30" height="20"><font class="form_corpo">
                                <%response.Write(f4)%>
                                </font></td>
                              <td width="450" height="20">&nbsp;</td>
                            </tr>
                          </table>
                        </div></td>
                    </tr>
                  </table>	
	<%
	end if
end if					
					%></td>
                      </tr>
                      <tr>
                        <td colspan="16" class="style1" >&nbsp;</td>
                      </tr>
                      <tr>
                        <td colspan="16" class="style1" >
                        	<% if libera_resultado="s" then
								observacao="Com base no Regimento Escolar e conforme publicado no informativo/2011, pág. 21.<BR>&nbsp;"
								if resultado_aluno = "APR" then
									resultado_aluno = "<B>Resultado: Aprovado</B><BR>&nbsp;<BR>"&observacao
								elseif resultado_aluno = "REC" then
									resultado_aluno = "<B>Resultado: Recuperação</B><BR>&nbsp;<BR>"&observacao
								elseif resultado_aluno = "PFI" then
									resultado_aluno = "<B>Resultado: Prova Final</B><BR>&nbsp;<BR>"&observacao
								elseif resultado_aluno = "REP" then
									resultado_aluno = "<B>Resultado: Reprovado</B><BR>&nbsp;<BR>"&observacao
								else
									resultado_aluno = "&nbsp;"
								end if
								response.Write(resultado_aluno)
							   else
								response.Write("&nbsp;")							   
							   end if
							%>
                        </td>
                      </tr>
                      <tr> 
                        <td colspan="16" 
>
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
                    </table>				
					</td>
                </tr>
              </table>
		</td>
  </tr>
  <tr>
    <td width="1000"><img src="../../img/rodape.jpg" width="1000" height="41" /></td>
  </tr>
</table>
</form>
</body>
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
response.redirect("../../inc/erro.asp")
end if
%>