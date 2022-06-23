<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->
<!--#include file="../../../../inc/bd_grade.asp"-->
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



		Set RS_turma = Server.CreateObject("ADODB.Recordset")
		SQL_turma = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND NU_Unidade ="& unidade &" AND CO_Curso='"&curso&"' AND CO_Etapa='"&etapa&"' AND CO_Turma='"&turma&"'"
		RS_turma.Open SQL_turma, CON1

		conta_alunos=0
		while not RS_turma.eof
		codigo_turma = RS_turma("CO_Matricula")
		
			if conta_alunos=0 then
			alunos_turma=codigo_turma
			else
			alunos_turma=codigo_turma&","&alunos_turma
			end if
		
		conta_alunos=conta_alunos+1
		RS_turma.MOVENEXT
		wend





ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

if mes<10 then
meswrt="0"&mes
else
meswrt=mes
end if
if min<10 then
minwrt="0"&min
else
minwrt=min
end if

data_proc = dia &"/"& meswrt &"/"& ano
horario = hora & ":"& minwrt


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

%></td>
  </tr>
  <tr>
    <td height="10" colspan="5" valign="top"><%call mensagens(nivel,636,0,0) %></td>
  </tr>
  <tr>
    <td valign="top"><table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
        <tr>
          <td width="653" class="tb_tit"
>Dados Escolares</td>
          <td width="113" class="tb_tit"
></td>
        </tr>
        <tr>
          <td height="10"><table width="100%" border="0" cellspacing="0">
              <tr>
                <td width="19%" height="10"><div align="right"><font class="form_dado_texto"> Matr&iacute;cula: </div></td>
                <td width="9%" height="10"><font class="form_corpo">
                  <input name="cod" type="hidden" value="<%=codigo%>">
                  <%response.Write(codigo)%></td>
                <td width="6%" height="10"><div align="right"><font class="form_dado_texto"> Nome: </div></td>
                <td width="66%" height="10"><font class="form_corpo">
                  <%response.Write(nome_prof)%>
                  <input name="nome" type="hidden" class="textInput" id="nome"  value="<%response.Write(nome_prof)%>" size="75" maxlength="50">
                  &nbsp;</td>
              </tr>
            </table></td>
          <td valign="top">&nbsp;</td>
        </tr>
        <tr>
          <td height="10" bgcolor="#FFFFFF">&nbsp;</td>
          <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
        <tr>
          <td colspan="2"><table width="100%" border="0" cellspacing="0">
              <tr class="tb_subtit">
                <td width="33" height="10"><div align="center">
                    <%
call GeraNomes("PORT",unidade,curso,etapa,CON0)
no_unidades = session("no_unidades")
no_grau = session("no_grau")
no_serie = session("no_serie")
%>
                    Ano</div></td>
                <td width="81" height="10"><div align="center">Matr&iacute;cula</div></td>
                <td width="75" height="10"><div align="center">Cancelamento</div></td>
                <td width="86" height="10"><div align="center"> Situa&ccedil;&atilde;o</div></td>
                <td width="113" height="10"><div align="center">Unidade</div></td>
                <td width="133" height="10"><div align="center">Curso</div></td>
                <td width="85" height="10"><div align="center"> Etapa</div></td>
                <td width="90" height="10"><div align="center">Turma </div></td>
                <td width="54" height="10"><div align="center">Chamada</div></td>
              </tr>
              <tr class="tb_corpo"
>
                <td width="33" height="10"><div align="center"> <font class="form_dado_texto">
                    <%response.Write(ano_aluno)%>
                  </div></td>
                <td width="81" height="10"><div align="center"> <font class="form_dado_texto">
                    <%response.Write(rematricula)%>
                  </div></td>
                <td width="75" height="10"><div align="center"> <font class="form_dado_texto">
                    <%response.Write(encerramento)%>
                  </div></td>
                <td width="86" height="10"><div align="center"> <font class="form_dado_texto">
                    <%
					
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
		RSCONTST.Open SQLCONTST, CON0
							
				no_situacao = RSCONTST("TX_Descricao_Situacao")	
					response.Write(no_situacao)%>
                  </div></td>
                <td width="113" height="10"><div align="center"> <font class="form_dado_texto">
                    <%response.Write(no_unidades)%>
                  </div></td>
                <td width="133" height="10"><div align="center"> <font class="form_dado_texto">
                    <%response.Write(no_grau)%>
                  </div></td>
                <td width="85" height="10"><div align="center"> <font class="form_dado_texto">
                    <%response.Write(no_serie)%>
                  </div></td>
                <td width="90" height="10"><div align="center"> <font class="form_dado_texto">
                    <%response.Write(turma)%>
                  </div></td>
                <td width="54" height="10"><div align="center"> <font class="form_dado_texto">
                    <%response.Write(cham)%>
                  </div></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td>&nbsp;</td>
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
                <td><%		Set RS_tb = Server.CreateObject("ADODB.Recordset")
		SQL_tb = "SELECT * FROM TB_Da_Aula WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso&"' AND CO_Etapa ='"& etapa&"' AND CO_Turma ='"& turma&"'"
		RS_tb.Open SQL_tb, CON2
if RS_tb.eof then
%>
                  <div align="center">
                    <%response.Write("<br><br><br><br><br>Não existe Boletim para este aluno!")%>
                  </div>
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
	
if session("ano_letivo") >= 2017 then
if notaFIL="TB_NOTA_A" or notaFIL="TB_NOTA_C" then			
	%>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="252" rowspan="2" class="tb_subtit"><div align="left"><strong>Disciplina</strong></div></td>
                      <td width="748" colspan="10" class="tb_subtit"><div align="center"></div>
                        <div align="center">Aproveitamento</div></td>
                    </tr>
                    <tr>
                      <td width="78" class="tb_subtit"><div align="center">TRI 
                          1</div></td>
                      <td width="78" class="tb_subtit"><div align="center">TRI 
                      2</div></td>
                      <td width="78" class="tb_subtit"><div align="center">TRI 
                      3</div></td>
                      <td width="78" class="tb_subtit"><div align="center">M&eacute;dia 
                          Anual</div></td>
                      <td width="78" class="tb_subtit"><div align="center">Result</div></td>
                      <td width="78" class="tb_subtit"><div align="center">Prova 
                          Final</div></td>
                      <td width="78" class="tb_subtit"><div align="center">M&eacute;dia 
                          Final</div></td>
                      <td width="78" class="tb_subtit"><div align="center">Result</div></td>
                      <td width="78" class="tb_subtit"><div align="center">Prova 
                          Recup</div></td>
                      <td width="78" class="tb_subtit"><div align="center">Result</div></td>
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
			
				Set RS1a = Server.CreateObject("ADODB.Recordset")
				SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
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
			
				for periodofil=1 to 5			
										
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
						'elseif periodofil=4 then
'						va_m34="&nbsp;"
'						va_m34_exibe="&nbsp;"
						elseif periodofil=4 then
						va_m35="&nbsp;"
						va_m35_exibe="&nbsp;"
						elseif periodofil=5 then
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
'						elseif periodofil=4 then
'						va_m34=RS3("VA_Media3")
'						va_m34_exibe=RS3("VA_Media3")
						elseif periodofil=4 then
						va_m35=RS3("VA_Media3")
						va_m35_exibe=RS3("VA_Media3")
						elseif periodofil=5 then
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
							if va_m31 >= 90 then
							va_m31_exibe="E"
							elseif (va_m31 > 79) and (va_m31 <= 89) then
							va_m31_exibe="MB"
							elseif (va_m31 > 69) and (va_m31 <= 79) then							
							va_m31_exibe="B"
							elseif (va_m31 > 59) and (va_m31 <= 69) then
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
							if va_m32 >= 90 then
									va_m32_exibe="E"
									elseif (va_m32 > 79) and (va_m32 <= 89) then
									va_m32_exibe="MB"
									elseif (va_m32 > 69) and (va_m32 <= 79) then							
									va_m32_exibe="B"
									elseif (va_m32 > 59) and (va_m32 <= 69) then
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
								if va_m33 >= 90 then
									va_m33_exibe="E"
									elseif (va_m33 > 79) and (va_m33 <= 89) then
									va_m33_exibe="MB"
									elseif (va_m33 > 69) and (va_m33 <= 79) then							
									va_m33_exibe="B"
									elseif (va_m33 > 59) and (va_m33 <= 69) then
									va_m33_exibe="R"
									else							
									va_m33_exibe="I"
								end if
						end if
					end if
					
'					if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
'					dividendo4=0
'					divisor4=0
'					else
'					dividendo4=va_m34
'					divisor4=1
'						if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then					
'								if va_m34 >= 90 then
'									va_m34_exibe="E"
'									elseif (va_m34 > 79) and (va_m34 <= 89) then
'									va_m34_exibe="MB"
'									elseif (va_m34 > 69) and (va_m34 <= 79) then							
'									va_m34_exibe="B"
'									elseif (va_m34 > 59) and (va_m34 <= 69) then
'									va_m34_exibe="R"
'									else							
'									va_m34_exibe="I"
'								end if
'						end if
'					end if
												
					dividendo_ma=dividendo1+dividendo2+dividendo3+dividendo4
					divisor_ma=divisor1+divisor2+divisor3+divisor4
					
					'response.Write(dividendo_ma&"<<")
					
					if divisor_ma<3 then
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
					
					'if ma="&nbsp;" then
					'dividendo_mf=0
					'divisor_mf=0
					'else
					'dividendo_mf=ma+dividendo5
					'divisor_mf=1+divisor5
					'end if
					
					'if divisor_mf=0 then
					'mf="&nbsp;"
					'else
					'response.Write(mf&"="&dividendo_mf&"/"&divisor_mf)
					'mf=dividendo_mf/divisor_mf
					'end if
					
					'if mf="&nbsp;" then
					'else
					'mf=mf/10
						'decimo = mf - Int(mf)
						'	If decimo >= 0.5 Then
						'		nota_arredondada = Int(mf) + 1
						'		mf=nota_arredondada
						'	Else
						'		nota_arredondada = Int(mf)
						'		mf=nota_arredondada					
						'	End If
						'mf = formatNumber(mf,1)
						'if mf>=minimo_recuperacao then
						'res2="APR"
						'else
						'res2="REC"
						'end if 						
					'end if	
					
					if isnull(va_m36) or va_m36="&nbsp;" or va_m36="" then
					nota_aux_m3_1="&nbsp;"
					'dividendo6=0
					'divisor6=0
					else
					nota_aux_m3_1=va_m36
					'dividendo6=va_m36
					'divisor6=1
					end if
					
					'if mf="&nbsp;" then
					'dividendo_rec=0
					'divisor_rec=0
					'else
					'dividendo_rec=mf+dividendo6
					'divisor_rec=1+divisor6
					'end if
					
					'if divisor_rec=0 then
					'rec="&nbsp;"
					'else
					'rec=dividendo_rec/divisor_rec
					'end if
					
					'if rec="&nbsp;" then
					'else
					'mf=mf/10
					'	decimo = rec - Int(rec)
					'		If decimo >= 0.5 Then
					'			nota_arredondada = Int(rec) + 1
					'			mf=nota_arredondada
					'		Else
					'			nota_arredondada = Int(rec)
					'			rec=nota_arredondada					
					'		End If
					'	rec = formatNumber(rec,1)

						'if rec>=minimo_aprovacao then
						'res3="APR"
						'else
						'res3="REP"
						'end if 							
					'end if				

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
							
							
									
					'end if
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
										if ma >= 90 then
										ma="E"
										elseif (ma > 79) and (ma <= 89) then
										ma="MB"
										elseif (ma > 69) and (ma <= 79) then							
										ma="B"
										elseif (ma > 59) and (ma <= 69) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 >= 90 then
										m2="E"
										elseif (m2 > 79) and (m2 <= 89) then
										m2="MB"
										elseif (m2 > 69) and (m2 <= 79) then							
										m2="B"
										elseif (m2 > 59) and (m2 <= 69) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if isnull(m3) or m3="" then
									'	m3=m2
									else
										if m3<>"&nbsp;" then									
											if m3 >= 90 then
											m3="E"
											elseif (m3 > 79) and (m3 <= 89) then
											m3="MB"
											elseif (m3 > 69) and (m3 <= 79) then							
											m3="B"
											elseif (m3 > 59) and (m3 <= 69) then
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
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%									if showapr1="s" then																	
									response.Write(va_m31_exibe)
									end if
							%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
							if showapr1="s" then												
							response.Write(va_m32_exibe)						
							end if
							%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
							if showapr1="s" then					
							response.Write(va_m33_exibe)
							end if
							%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
							if showapr1="s" then
							response.Write(ma)
							else
							end if
							%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
							if showapr2="s" and showprova2="s" then
																				
							response.Write(res1)					
							end if
							%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
							if showapr2="s" and showprova2="s" then												
							response.Write(va_m35_exibe)
							end if
							%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
							if showapr2="s" and showprova2="s" then					
							response.Write(m2)
							end if
							%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
							if showapr2="s" and showprova2="s" then					
							response.Write(res2)
							end if
							%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
							if showapr2="s" and showprova2="s" then
							response.Write(m3)
							else
							end if
							%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
							if showapr3="s" and showprova3="s" then													
							response.Write(res3)	
							end if

							%>
                        </div></td>
                    </tr>
                    <%
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
                      <td height="20" colspan="11"><div align="right">
                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td width="240"><font class="form_dado_texto">Freq&uuml;&ecirc;ncia 
                                (Faltas):</font></td>
                              <td width="77"><div align="right"><font class="form_dado_texto">Trimestre 
                                  1:</font></div></td>
                              <td width="29"><font class="form_corpo">
                                <%response.Write(f1)%>
                                </font></td>
                              <td width="69"><div align="right"><font class="form_dado_texto">Trimestre 
                                  2:</font></div></td>
                              <td width="29"><font class="form_corpo">
                                <%response.Write(f2)%>
                                </font></td>
                              <td width="69"><div align="right"><font class="form_dado_texto">Trimestre 
                                  3:</font></div></td>
                              <td width="29"><font class="form_corpo">
                                <%response.Write(f3)%>
                                </font></td>
                              <td width="454">&nbsp;</td>
                            </tr>
                          </table>
                      </div></td>
                    </tr>
                  </table>
                  <%
	elseif notaFIL="TB_NOTA_B" or notaFIL="TB_NOTA_E" or notaFIL="TB_NOTA_F" or notaFIL="TB_NOTA_K" or notaFIL="TB_NOTA_L" or notaFIL="TB_NOTA_M" then
	%>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="252" rowspan="2" class="tb_subtit"><div align="left"><strong>Disciplina</strong></div></td>
                      <td width="748" colspan="10" class="tb_subtit"><div align="center"></div>
                        <div align="center">Aproveitamento</div></td>
                    </tr>
                    <tr>
                      <td width="78" class="tb_subtit"><div align="center">TRI 
                      1</div></td>
                      <td width="78" class="tb_subtit"><div align="center">TRI 
                      2</div></td>
                      <td width="78" class="tb_subtit"><div align="center">TRI 
                      3</div></td>
                      <td width="78" class="tb_subtit"><div align="center">M&eacute;dia 
                          Anual</div></td>
                      <td width="78" class="tb_subtit"><div align="center">Result</div></td>
                      <td width="78" class="tb_subtit"><div align="center">Prova 
                          Final</div></td>
                      <td width="78" class="tb_subtit"><div align="center">M&eacute;dia 
                          Final</div></td>
                      <td width="78" class="tb_subtit"><div align="center">Result</div></td>
                      <td width="78" class="tb_subtit"><div align="center">Prova 
                          Recup</div></td>
                      <td width="78" class="tb_subtit"><div align="center">Result</div></td>
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
					
						for periodofil=1 to 5
					
								
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
'						elseif periodofil=4 then
'						va_m34="&nbsp;"
'						va_m34_exibe="&nbsp;"
						elseif periodofil=4 then
						va_m35="&nbsp;"
						va_m35_exibe="&nbsp;"
						elseif periodofil=5 then
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
'						elseif periodofil=4 then
'						va_m34=RS3("VA_Media3")
'						va_m34_exibe=RS3("VA_Media3")
						elseif periodofil=4 then
						va_m35=RS3("VA_Media3")
						va_m35_exibe=RS3("VA_Media3")
						elseif periodofil=5 then
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
																if va_m31 >= 90 then
							va_m31_exibe="E"
							elseif (va_m31 > 79) and (va_m31 <= 89) then
							va_m31_exibe="MB"
							elseif (va_m31 > 69) and (va_m31 <= 79) then							
							va_m31_exibe="B"
							elseif (va_m31 > 59) and (va_m31 <= 69) then
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
														if va_m32 >= 90 then
									va_m32_exibe="E"
									elseif (va_m32 > 79) and (va_m32 <= 89) then
									va_m32_exibe="MB"
									elseif (va_m32 > 69) and (va_m32 <= 79) then							
									va_m32_exibe="B"
									elseif (va_m32 > 59) and (va_m32 <= 69) then
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
								if va_m33 >= 90 then
									va_m33_exibe="E"
									elseif (va_m33 > 79) and (va_m33 <= 89) then
									va_m33_exibe="MB"
									elseif (va_m33 > 69) and (va_m33 <= 79) then							
									va_m33_exibe="B"
									elseif (va_m33 > 59) and (va_m33 <= 69) then
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
								if va_m34 >= 90 then
									va_m34_exibe="E"
									elseif (va_m34 > 79) and (va_m34 <= 89) then
									va_m34_exibe="MB"
									elseif (va_m34 > 69) and (va_m34 <= 79) then							
									va_m34_exibe="B"
									elseif (va_m34 > 59) and (va_m34 <= 69) then
									va_m34_exibe="R"
									else							
									va_m34_exibe="I"
								end if
						end if
							end if
										
							dividendo_ma=dividendo1+dividendo2+dividendo3+dividendo4
							divisor_ma=divisor1+divisor2+divisor3+divisor4
							
											
							if divisor_ma<3 then
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
										if ma >= 90 then
										ma="E"
										elseif (ma > 79) and (ma <= 89) then
										ma="MB"
										elseif (ma > 69) and (ma <= 79) then							
										ma="B"
										elseif (ma > 59) and (ma <= 69) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 >= 90 then
										m2="E"
										elseif (m2 > 79) and (m2 <= 89) then
										m2="MB"
										elseif (m2 > 69) and (m2 <= 79) then							
										m2="B"
										elseif (m2 > 59) and (m2 <= 69) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if isnull(m3) or m3="" then
									'	m3=m2
									else
										if m3<>"&nbsp;" then									
											if m3 >= 90 then
											m3="E"
											elseif (m3 > 79) and (m3 <= 89) then
											m3="MB"
											elseif (m3 > 69) and (m3 <= 79) then							
											m3="B"
											elseif (m3 > 59) and (m3 <= 69) then
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
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then																	
									response.Write(va_m31_exibe)
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then												
									response.Write(va_m32_exibe)						
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then					
									response.Write(va_m33_exibe)
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then
									response.Write(ma)
									else
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then
																						
									response.Write(res1)					
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then												
									response.Write(va_m35_exibe)
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then					
									response.Write(m2)
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then					
									response.Write(res2)
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then
									response.Write(m3)
									else
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
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
					
						for periodofil=1 to 5
					
								
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
'						elseif periodofil=4 then
'						va_m34="&nbsp;"
'						va_m34_exibe="&nbsp;"
						elseif periodofil=4 then
						va_m35="&nbsp;"
						va_m35_exibe="&nbsp;"
						elseif periodofil=5 then
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
'						va_m34=RS3("VA_Media3")
'						va_m34_exibe=RS3("VA_Media3")
						elseif periodofil=4 then
						va_m35=RS3("VA_Media3")
						va_m35_exibe=RS3("VA_Media3")
						elseif periodofil=5 then
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
								if va_m31 >= 90 then
									va_m31_exibe="E"
								elseif (va_m31 > 79) and (va_m31 <= 89) then
									va_m31_exibe="MB"
								elseif (va_m31 > 69) and (va_m31 <= 79) then							
									va_m31_exibe="B"
								elseif (va_m31 > 59) and (va_m31 <= 69) then
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
														if va_m32 >= 90 then
									va_m32_exibe="E"
									elseif (va_m32 > 79) and (va_m32 <= 89) then
									va_m32_exibe="MB"
									elseif (va_m32 > 69) and (va_m32 <= 79) then							
									va_m32_exibe="B"
									elseif (va_m32 > 59) and (va_m32 <= 69) then
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
								if va_m33 >= 90 then
									va_m33_exibe="E"
									elseif (va_m33 > 79) and (va_m33 <= 89) then
									va_m33_exibe="MB"
									elseif (va_m33 > 69) and (va_m33 <= 79) then							
									va_m33_exibe="B"
									elseif (va_m33 > 59) and (va_m33 <= 69) then
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
								if va_m34 >= 90 then
									va_m34_exibe="E"
									elseif (va_m34 > 79) and (va_m34 <= 89) then
									va_m34_exibe="MB"
									elseif (va_m34 > 69) and (va_m34 <= 79) then							
									va_m34_exibe="B"
									elseif (va_m34 > 59) and (va_m34 <= 69) then
									va_m34_exibe="R"
									else							
									va_m34_exibe="I"
								end if
						end if
							end if
										
							dividendo_ma=dividendo1+dividendo2+dividendo3+dividendo4
							divisor_ma=divisor1+divisor2+divisor3+divisor4
											
							if divisor_ma<3 then
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
										if ma >= 90 then
										ma="E"
										elseif (ma > 79) and (ma <= 89) then
										ma="MB"
										elseif (ma > 69) and (ma <= 79) then							
										ma="B"
										elseif (ma > 59) and (ma <= 69) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 >= 90 then
										m2="E"
										elseif (m2 > 79) and (m2 <= 89) then
										m2="MB"
										elseif (m2 > 69) and (m2 <= 79) then							
										m2="B"
										elseif (m2 > 59) and (m2 <= 69) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if isnull(m3) or m3="" then
									'	m3=m2
									else
										if m3<>"&nbsp;" then									
											if m3 >= 90 then
											m3="E"
											elseif (m3 > 79) and (m3 <= 89) then
											m3="MB"
											elseif (m3 > 69) and (m3 <= 79) then							
											m3="B"
											elseif (m3 > 59) and (m3 <= 69) then
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
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%									if showapr1="s" then																	
									response.Write(va_m31_exibe)
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then												
									response.Write(va_m32_exibe)						
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then					
									response.Write(va_m33_exibe)
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then
									response.Write(ma)
									else
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then
																						
									response.Write(res1)					
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then												
									response.Write(va_m35_exibe)
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then					
									response.Write(m2)
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then					
									response.Write(res2)
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then
									response.Write(m3)
									else
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
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
							
							for periodofil=1 to 5	
										
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
'						elseif periodofil=4 then
'						va_m34="&nbsp;"
'						va_m34_exibe="&nbsp;"
'						conta4=conta4
						elseif periodofil=4 then
						va_m35="&nbsp;"
						va_m35_exibe="&nbsp;"
						conta5=conta5
						elseif periodofil=5 then
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
'						elseif periodofil=4 then
'						va_m34=RS3("VA_Media3")
'						va_m34_exibe=RS3("VA_Media3")
'								if isnull(va_m34_exibe) or va_m34_exibe="" then
'								conta4=conta4
'								else
'								conta4=conta4+1
'								end if						
						elseif periodofil=4 then
						va_m35=RS3("VA_Media3")
						va_m35_exibe=RS3("VA_Media3")
								if isnull(va_m35_exibe) or va_m35_exibe="" then
								conta5=conta5
								else
								conta5=conta5+1
								end if						
						elseif periodofil=5 then
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
																	if va_m31 >= 90 then
								va_m31_exibe="E"
								elseif (va_m31 > 79) and (va_m31 <= 89) then
								va_m31_exibe="MB"
								elseif (va_m31 > 69) and (va_m31 <= 79) then							
								va_m31_exibe="B"
								elseif (va_m31 > 59) and (va_m31 <= 69) then
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
							
								if va_m32 >= 90 then
									va_m32_exibe="E"
								elseif (va_m32 > 79) and (va_m32 <= 89) then
									va_m32_exibe="MB"
								elseif (va_m32 > 69) and (va_m32 <= 79) then							
									va_m32_exibe="B"
								elseif (va_m32 > 59) and (va_m32 <= 69) then
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
									if va_m33 >= 90 then
									va_m33_exibe="E"
									elseif (va_m33 > 79) and (va_m33 <= 89) then
									va_m33_exibe="MB"
									elseif (va_m33 > 69) and (va_m33 <= 79) then							
									va_m33_exibe="B"
									elseif (va_m33 > 59) and (va_m33 <= 69) then
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
									if va_m34 >= 90 then
									va_m34_exibe="E"
									elseif (va_m34 > 79) and (va_m34 <= 89) then
									va_m34_exibe="MB"
									elseif (va_m34 > 69) and (va_m34 <= 79) then							
									va_m34_exibe="B"
									elseif (va_m34 > 59) and (va_m34 <= 69) then
									va_m34_exibe="R"
									else							
									va_m34_exibe="I"
								end if
								end if
							end if
										
							dividendo_ma=dividendo1+dividendo2+dividendo3+dividendo4
							divisor_ma=divisor1+divisor2+divisor3+divisor4
											
							if divisor_ma<3 then
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
								resultados=novo_regra_aprovacao (cod, materia,curso,etapa,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,"waboletim")
'response.Write(materia&":"&resultados&"<BR>")									
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
										if ma >= 90 then
										ma="E"
										elseif (ma > 79) and (ma <= 89) then
										ma="MB"
										elseif (ma > 69) and (ma <= 79) then							
										ma="B"
										elseif (ma > 59) and (ma <= 69) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 >= 90 then
										m2="E"
										elseif (m2 > 79) and (m2 <= 89) then
										m2="MB"
										elseif (m2 > 69) and (m2 <= 79) then							
										m2="B"
										elseif (m2 > 59) and (m2 <= 69) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if isnull(m3) or m3="" then
									'	m3=m2
									else
										if m3<>"&nbsp;" then									
											if m3 >= 90 then
											m3="E"
											elseif (m3 > 79) and (m3 <= 89) then
											m3="MB"
											elseif (m3 > 69) and (m3 <= 79) then							
											m3="B"
											elseif (m3 > 59) and (m3 <= 69) then
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
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%									if showapr1="s" then																	
									response.Write(va_m31_exibe)
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then												
									response.Write(va_m32_exibe)						
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then					
									response.Write(va_m33_exibe)
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr1="s" then
									response.Write(ma)
									else
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then
																						
									'response.Write(res1)					
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then												
									response.Write(va_m35_exibe)
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then	
										if notaFIL="TB_NOTA_L" then
											response.Write(m2)										
										else
											response.Write(va_m35)
										end if	
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then					
									'response.Write(res2)
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
                          <%							
									if showapr2="s" and showprova2="s" then
										if notaFIL="TB_NOTA_L" then
											response.Write(m3)										
										else
											response.Write(va_m36)
										end if
									else
									end if
									%>
                        </div></td>
                      <td width="78" class="<%response.Write(cor)%>"><div align="center">
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
										
							minimo_exibir=3
							'response.Write(va_m31&" - "&va_m32&" - "&va_m33&" - "&va_m34&" - "&divisor_m_acumul&"<"&minimo_exibir)								
							if divisor_m_acumul<minimo_exibir then
								m_acumul="&nbsp;"
							else
								m31_mae=m31_mae*1
								m32_mae=m32_mae*1
								m33_mae=m33_mae*1
								m34_mae=0
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
										if ma >= 90 then
										ma="E"
										elseif (ma > 79) and (ma <= 89) then
										ma="MB"
										elseif (ma > 69) and (ma <= 79) then							
										ma="B"
										elseif (ma > 59) and (ma <= 69) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 >= 90 then
										m2="E"
										elseif (m2 > 79) and (m2 <= 89) then
										m2="MB"
										elseif (m2 > 69) and (m2 <= 79) then							
										m2="B"
										elseif (m2 > 59) and (m2 <= 69) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if isnull(m3) or m3="" then
									'	m3=m2
									else
										if m3<>"&nbsp;" then									
											if m3 >= 90 then
											m3="E"
											elseif (m3 > 79) and (m3 <= 89) then
											m3="MB"
											elseif (m3 > 69) and (m3 <= 79) then							
											m3="B"
											elseif (m3 > 59) and (m3 <= 69) then
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
                      <td width="78" ><div align="center">
                          <%
									if showapr1="s" then																	
									response.Write(m31_exibe)
									end if
									%>
                        </div></td>
                      <td width="78" ><div align="center">
                          <%
									if showapr1="s" then												
									response.Write(m32_exibe)						
									end if
									%>
                        </div></td>
                      <td width="78" ><div align="center">
                          <%
									if showapr1="s" then					
									response.Write(m33_exibe)
									end if
									%>
                        </div></td>
                      <td width="78" ><div align="center">
                          <%
									if showapr1="s" then
									response.Write(m_acumul)
									else
									end if
									%>
                        </div></td>
                      <td width="78" ><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then
																						
									response.Write(res1)					
									end if
									%>
                        </div></td>
                      <td width="78" ><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then												
									response.Write(m35_mae)
									end if
									%>
                        </div></td>
                      <td width="78" ><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then					
									response.Write(m2)
									end if
									%>
                        </div></td>
                      <td width="78" ><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then					
									response.Write(res2)
									end if
									%>
                        </div></td>
                      <td width="78" ><div align="center">
                          <%
									if showapr2="s" and showprova2="s" then
									response.Write(m3)
									else
									end if
									%>
                        </div></td>
                      <td width="78" ><div align="center">
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
                      <td height="20" colspan="11"><div align="right">
                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr valign="middle">
                              <td width="250" height="20"><font class="form_dado_texto">Freq&uuml;&ecirc;ncia 
                                (Faltas):</font></td>
                              <td width="70" height="20"><div align="right"><font class="form_dado_texto">Trimestre 
                              1:</font></div></td>
                              <td width="30" height="20"><font class="form_corpo">
                                <%response.Write(f1)%>
                                </font></td>
                              <td width="70" height="20"><div align="right"><font class="form_dado_texto">Trimestre 
                              2:</font></div></td>
                              <td width="30" height="20"><font class="form_corpo">
                                <%response.Write(f2)%>
                                </font></td>
                              <td width="70" height="20"><div align="right"><font class="form_dado_texto">Trimestre 
                              3:</font></div></td>
                              <td width="30" height="20"><font class="form_corpo">
                                <%response.Write(f3)%>
                                </font></td>
                              <td width="450" height="20">&nbsp;</td>
                            </tr>
                          </table>
                      </div></td>
                    </tr>
                  </table>
<%
	end if
else
if notaFIL="TB_NOTA_A" or notaFIL="TB_NOTA_C" then			
	%>
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
			
				Set RS1a = Server.CreateObject("ADODB.Recordset")
				SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
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
							if va_m31 >= 90 then
							va_m31_exibe="E"
							elseif (va_m31 > 79) and (va_m31 <= 89) then
							va_m31_exibe="MB"
							elseif (va_m31 > 69) and (va_m31 <= 79) then							
							va_m31_exibe="B"
							elseif (va_m31 > 59) and (va_m31 <= 69) then
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
							if va_m32 >= 90 then
									va_m32_exibe="E"
									elseif (va_m32 > 79) and (va_m32 <= 89) then
									va_m32_exibe="MB"
									elseif (va_m32 > 69) and (va_m32 <= 79) then							
									va_m32_exibe="B"
									elseif (va_m32 > 59) and (va_m32 <= 69) then
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
								if va_m33 >= 90 then
									va_m33_exibe="E"
									elseif (va_m33 > 79) and (va_m33 <= 89) then
									va_m33_exibe="MB"
									elseif (va_m33 > 69) and (va_m33 <= 79) then							
									va_m33_exibe="B"
									elseif (va_m33 > 59) and (va_m33 <= 69) then
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
								if va_m34 >= 90 then
									va_m34_exibe="E"
									elseif (va_m34 > 79) and (va_m34 <= 89) then
									va_m34_exibe="MB"
									elseif (va_m34 > 69) and (va_m34 <= 79) then							
									va_m34_exibe="B"
									elseif (va_m34 > 59) and (va_m34 <= 69) then
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
					
					'if ma="&nbsp;" then
					'dividendo_mf=0
					'divisor_mf=0
					'else
					'dividendo_mf=ma+dividendo5
					'divisor_mf=1+divisor5
					'end if
					
					'if divisor_mf=0 then
					'mf="&nbsp;"
					'else
					'response.Write(mf&"="&dividendo_mf&"/"&divisor_mf)
					'mf=dividendo_mf/divisor_mf
					'end if
					
					'if mf="&nbsp;" then
					'else
					'mf=mf/10
						'decimo = mf - Int(mf)
						'	If decimo >= 0.5 Then
						'		nota_arredondada = Int(mf) + 1
						'		mf=nota_arredondada
						'	Else
						'		nota_arredondada = Int(mf)
						'		mf=nota_arredondada					
						'	End If
						'mf = formatNumber(mf,1)
						'if mf>=minimo_recuperacao then
						'res2="APR"
						'else
						'res2="REC"
						'end if 						
					'end if	
					
					if isnull(va_m36) or va_m36="&nbsp;" or va_m36="" then
					nota_aux_m3_1="&nbsp;"
					'dividendo6=0
					'divisor6=0
					else
					nota_aux_m3_1=va_m36
					'dividendo6=va_m36
					'divisor6=1
					end if
					
					'if mf="&nbsp;" then
					'dividendo_rec=0
					'divisor_rec=0
					'else
					'dividendo_rec=mf+dividendo6
					'divisor_rec=1+divisor6
					'end if
					
					'if divisor_rec=0 then
					'rec="&nbsp;"
					'else
					'rec=dividendo_rec/divisor_rec
					'end if
					
					'if rec="&nbsp;" then
					'else
					'mf=mf/10
					'	decimo = rec - Int(rec)
					'		If decimo >= 0.5 Then
					'			nota_arredondada = Int(rec) + 1
					'			mf=nota_arredondada
					'		Else
					'			nota_arredondada = Int(rec)
					'			rec=nota_arredondada					
					'		End If
					'	rec = formatNumber(rec,1)

						'if rec>=minimo_aprovacao then
						'res3="APR"
						'else
						'res3="REP"
						'end if 							
					'end if				

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
							
							
									
					'end if
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
										if ma >= 90 then
										ma="E"
										elseif (ma > 79) and (ma <= 89) then
										ma="MB"
										elseif (ma > 69) and (ma <= 79) then							
										ma="B"
										elseif (ma > 59) and (ma <= 69) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 >= 90 then
										m2="E"
										elseif (m2 > 79) and (m2 <= 89) then
										m2="MB"
										elseif (m2 > 69) and (m2 <= 79) then							
										m2="B"
										elseif (m2 > 59) and (m2 <= 69) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if isnull(m3) or m3="" then
									'	m3=m2
									else
										if m3<>"&nbsp;" then									
											if m3 >= 90 then
											m3="E"
											elseif (m3 > 79) and (m3 <= 89) then
											m3="MB"
											elseif (m3 > 69) and (m3 <= 79) then							
											m3="B"
											elseif (m3 > 59) and (m3 <= 69) then
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
                            <tr>
                              <td width="250"><font class="form_dado_texto">Freq&uuml;&ecirc;ncia 
                                (Faltas):</font></td>
                              <td width="70"><div align="right"><font class="form_dado_texto">Bimestre 
                                  1:</font></div></td>
                              <td width="30"><font class="form_corpo">
                                <%response.Write(f1)%>
                                </font></td>
                              <td width="70"><div align="right"><font class="form_dado_texto">Bimestre 
                                  2:</font></div></td>
                              <td width="30"><font class="form_corpo">
                                <%response.Write(f2)%>
                                </font></td>
                              <td width="70"><div align="right"><font class="form_dado_texto">Bimestre 
                                  3:</font></div></td>
                              <td width="30"><font class="form_corpo">
                                <%response.Write(f3)%>
                                </font></td>
                              <td width="70"><div align="right"><font class="form_dado_texto">Bimestre 
                                  4:</font></div></td>
                              <td width="30"><font class="form_corpo">
                                <%response.Write(f4)%>
                                </font></td>
                              <td width="450">&nbsp;</td>
                            </tr>
                          </table>
                        </div></td>
                    </tr>
                  </table>
                  <%
	elseif notaFIL="TB_NOTA_B" or notaFIL="TB_NOTA_E" or notaFIL="TB_NOTA_F" or notaFIL="TB_NOTA_K" or notaFIL="TB_NOTA_L" or notaFIL="TB_NOTA_M" then
	%>
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
									if va_m31 >= 90 then
									va_m31_exibe="E"
									elseif (va_m31 > 79) and (va_m31 <= 89) then
									va_m31_exibe="MB"
									elseif (va_m31 > 69) and (va_m31 <= 79) then							
									va_m31_exibe="B"
									elseif (va_m31 > 59) and (va_m31 <= 69) then
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
									if va_m32 >= 90 then
									va_m32_exibe="E"
									elseif (va_m32 > 79) and (va_m32 <= 89) then
									va_m32_exibe="MB"
									elseif (va_m32 > 69) and (va_m32 <= 79) then							
									va_m32_exibe="B"
									elseif (va_m32 > 59) and (va_m32 <= 69) then
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
									if va_m33 >= 90 then
									va_m33_exibe="E"
									elseif (va_m33 > 79) and (va_m33 <= 89) then
									va_m33_exibe="MB"
									elseif (va_m33 > 69) and (va_m33 <= 79) then							
									va_m33_exibe="B"
									elseif (va_m33 > 59) and (va_m33 <= 69) then
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
									if va_m34 >= 90 then
										va_m34_exibe="E"
										elseif (va_m34 > 79) and (va_m34 <= 89) then
										va_m34_exibe="MB"
										elseif (va_m34 > 69) and (va_m34 <= 79) then							
										va_m34_exibe="B"
										elseif (va_m34 > 59) and (va_m34 <= 69) then
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
										if ma >= 90 then
										ma="E"
										elseif (ma > 79) and (ma <= 89) then
										ma="MB"
										elseif (ma > 69) and (ma <= 79) then							
										ma="B"
										elseif (ma > 59) and (ma <= 69) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 >= 90 then
										m2="E"
										elseif (m2 > 79) and (m2 <= 89) then
										m2="MB"
										elseif (m2 > 69) and (m2 <= 79) then							
										m2="B"
										elseif (m2 > 59) and (m2 <= 69) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if isnull(m3) or m3="" then
									'	m3=m2
									else
										if m3<>"&nbsp;" then									
											if m3 >= 90 then
											m3="E"
											elseif (m3 > 79) and (m3 <= 89) then
											m3="MB"
											elseif (m3 > 69) and (m3 <= 79) then							
											m3="B"
											elseif (m3 > 59) and (m3 <= 69) then
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
								if va_m31 >= 90 then
									va_m31_exibe="E"
								elseif (va_m31 > 79) and (va_m31 <= 89) then
									va_m31_exibe="MB"
								elseif (va_m31 > 69) and (va_m31 <= 79) then							
									va_m31_exibe="B"
								elseif (va_m31 > 59) and (va_m31 <= 69) then
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
														if va_m32 >= 90 then
									va_m32_exibe="E"
									elseif (va_m32 > 79) and (va_m32 <= 89) then
									va_m32_exibe="MB"
									elseif (va_m32 > 69) and (va_m32 <= 79) then							
									va_m32_exibe="B"
									elseif (va_m32 > 59) and (va_m32 <= 69) then
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
								if va_m33 >= 90 then
									va_m33_exibe="E"
									elseif (va_m33 > 79) and (va_m33 <= 89) then
									va_m33_exibe="MB"
									elseif (va_m33 > 69) and (va_m33 <= 79) then							
									va_m33_exibe="B"
									elseif (va_m33 > 59) and (va_m33 <= 69) then
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
								if va_m34 >= 90 then
									va_m34_exibe="E"
									elseif (va_m34 > 79) and (va_m34 <= 89) then
									va_m34_exibe="MB"
									elseif (va_m34 > 69) and (va_m34 <= 79) then							
									va_m34_exibe="B"
									elseif (va_m34 > 59) and (va_m34 <= 69) then
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
										if ma >= 90 then
										ma="E"
										elseif (ma > 79) and (ma <= 89) then
										ma="MB"
										elseif (ma > 69) and (ma <= 79) then							
										ma="B"
										elseif (ma > 59) and (ma <= 69) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 >= 90 then
										m2="E"
										elseif (m2 > 79) and (m2 <= 89) then
										m2="MB"
										elseif (m2 > 69) and (m2 <= 79) then							
										m2="B"
										elseif (m2 > 59) and (m2 <= 69) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if isnull(m3) or m3="" then
									'	m3=m2
									else
										if m3<>"&nbsp;" then									
											if m3 >= 90 then
											m3="E"
											elseif (m3 > 79) and (m3 <= 89) then
											m3="MB"
											elseif (m3 > 69) and (m3 <= 79) then							
											m3="B"
											elseif (m3 > 59) and (m3 <= 69) then
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
																	if va_m31 >= 90 then
								va_m31_exibe="E"
								elseif (va_m31 > 79) and (va_m31 <= 89) then
								va_m31_exibe="MB"
								elseif (va_m31 > 69) and (va_m31 <= 79) then							
								va_m31_exibe="B"
								elseif (va_m31 > 59) and (va_m31 <= 69) then
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
							
								if va_m32 >= 90 then
									va_m32_exibe="E"
								elseif (va_m32 > 79) and (va_m32 <= 89) then
									va_m32_exibe="MB"
								elseif (va_m32 > 69) and (va_m32 <= 79) then							
									va_m32_exibe="B"
								elseif (va_m32 > 59) and (va_m32 <= 69) then
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
									if va_m33 >= 90 then
									va_m33_exibe="E"
									elseif (va_m33 > 79) and (va_m33 <= 89) then
									va_m33_exibe="MB"
									elseif (va_m33 > 69) and (va_m33 <= 79) then							
									va_m33_exibe="B"
									elseif (va_m33 > 59) and (va_m33 <= 69) then
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
									if va_m34 >= 90 then
									va_m34_exibe="E"
									elseif (va_m34 > 79) and (va_m34 <= 89) then
									va_m34_exibe="MB"
									elseif (va_m34 > 69) and (va_m34 <= 79) then							
									va_m34_exibe="B"
									elseif (va_m34 > 59) and (va_m34 <= 69) then
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
										if ma >= 90 then
										ma="E"
										elseif (ma > 79) and (ma <= 89) then
										ma="MB"
										elseif (ma > 69) and (ma <= 79) then							
										ma="B"
										elseif (ma > 59) and (ma <= 69) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 >= 90 then
										m2="E"
										elseif (m2 > 79) and (m2 <= 89) then
										m2="MB"
										elseif (m2 > 69) and (m2 <= 79) then							
										m2="B"
										elseif (m2 > 59) and (m2 <= 69) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if isnull(m3) or m3="" then
									'	m3=m2
									else
										if m3<>"&nbsp;" then									
											if m3 >= 90 then
											m3="E"
											elseif (m3 > 79) and (m3 <= 89) then
											m3="MB"
											elseif (m3 > 69) and (m3 <= 79) then							
											m3="B"
											elseif (m3 > 59) and (m3 <= 69) then
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
										if ma >= 90 then
										ma="E"
										elseif (ma > 79) and (ma <= 89) then
										ma="MB"
										elseif (ma > 69) and (ma <= 79) then							
										ma="B"
										elseif (ma > 59) and (ma <= 69) then
										ma="R"
										else							
										ma="I"
										end if
									end if	
									if m2<>"&nbsp;" then
										if m2 >= 90 then
										m2="E"
										elseif (m2 > 79) and (m2 <= 89) then
										m2="MB"
										elseif (m2 > 69) and (m2 <= 79) then							
										m2="B"
										elseif (m2 > 59) and (m2 <= 69) then
										m2="R"
										else							
										m2="I"
										end if
									end if	
									if isnull(m3) or m3="" then
									'	m3=m2
									else
										if m3<>"&nbsp;" then									
											if m3 >= 90 then
											m3="E"
											elseif (m3 > 79) and (m3 <= 89) then
											m3="MB"
											elseif (m3 > 69) and (m3 <= 79) then							
											m3="B"
											elseif (m3 > 59) and (m3 <= 69) then
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
end if					
					%></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
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