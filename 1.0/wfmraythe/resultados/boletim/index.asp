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

		Set RSapr = Server.CreateObject("ADODB.Recordset")
		SQLapr = "Select * from TB_Regras_Aprovacao WHERE CO_Curso = '"& curso &"' AND CO_Etapa='"&etapa&"'"
		Set RSapr = CON0.Execute(SQLapr)
		
		if RSapr.EOF then
			ntvml=0
		else
			ntazl= RSapr("NU_Valor_M1")		
			ntvml= RSapr("NU_Valor_M2")
		end if
		cor_nota_vml="#FF0000"	
		cor_nota_azl="#0000FF"	
		cor_nota_prt="#000000"	

%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Web Fam�lia</title>
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
<form action="index.asp?opt=1" method="post">
  <table width="1000" height="1039" border="0" align="center" cellpadding="0" cellspacing="0" class="tabela_principal">
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
 mes = "mar�o"
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
%>
                    <div align="center"><font class="style1"> 
                      <%response.Write("<br><br><br><br><br>N�o existe Boletim para este aluno!")%>
                      </font> </div>
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
	%>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="339" height="20" class="style3"> 
                          <div align="left"></div></td>
                        <td height="20" colspan="7" class="tb_tit"> 
                          <div align="center"></div>
                          <div align="center">Aproveitamento</div></td>
                        <td width="1" height="20"> 
                          <div align="center"></div></td>
                        <td height="20" colspan="4" class="tb_tit"> 
                          <div align="center">Freq&uuml;&ecirc;ncia 
                            (Faltas):</div></td>
                      </tr>
                      <tr>
                        <td width="339" height="20" class="style3"><strong>Disciplina</strong></td>
                        <td width="60" height="20" class="style3"> 
                          <div align="center">PA1</div></td>
                        <td width="60" height="20" class="style3"> 
                          <div align="center">PA2</div></td>
                        <td width="60" height="20" class="style3"> 
                          <div align="center">PA3</div></td>
                        <td width="60" height="20" class="style3"> 
                          <div align="center">TOTAL</div></td>
                        <td width="60" height="20" class="style3"> 
                          <div align="center">4&ordf; 
                            aval<br>
                            p.2</div></td>
                        <td width="60" height="20" class="style3"> 
                          <div align="center">TOTAL</div></td>
                        <td width="60" height="20" class="style3"> 
                          <div align="center">M&eacute;dia 
                            Final</div></td>
                        <td width="1" height="20">&nbsp;</td>
                        <td width="60" height="20" class="style3"> 
                          <div align="center">PA1</div></td>
                        <td width="60" height="20" class="style3"> 
                          <div align="center">PA2</div></td>
                        <td width="60" height="20" class="style3"> 
                          <div align="center">PA3</div></td>
                        <td width="60" height="20" class="style3"> 
                          <div align="center">TOTAL</div></td>
                      </tr>
                      <%
			rec_lancado="sim"
		
			Set RSprog = Server.CreateObject("ADODB.Recordset")
			SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' AND IN_MAE=TRUE order by NU_Ordem_Boletim "
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
			
				if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(nu_peso)) or (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(nu_peso)) or (mae=TRUE and fil=TRUE and in_co=FALSE) then				
			
					for periodofil=1 to 4
					
						Set RS3 = Server.CreateObject("ADODB.Recordset")
						SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' AND CO_Materia_Principal = '"& materia &"' AND NU_Periodo="&periodofil
						Set RS3 = CON_N.Execute(SQL_N)
					
					
					
						if RS3.EOF then
							if periodofil=1 then
							f1="&nbsp;"
							va_m31="&nbsp;"
							elseif periodofil=2 then
							f2="&nbsp;"
							va_m32="&nbsp;"
							elseif periodofil=3 then
							f3="&nbsp;"
							va_m33="&nbsp;"
							elseif periodofil=4 then
							f4="&nbsp;"
							va_m34="&nbsp;"
							end if	
						else
							if periodofil=1 then
								f1=RS3("NU_Faltas")						
								va_m31=RS3("VA_Media3")
							elseif periodofil=2 then
								f2=RS3("NU_Faltas")						
								va_m32=RS3("VA_Media3")
							elseif periodofil=3 then
								f3=RS3("NU_Faltas")						
								va_m33=RS3("VA_Media3")
							elseif periodofil=4 then
								f4=RS3("NU_Faltas")
								va_m34=RS3("VA_Media3")
							end if
						end if
					NEXT		
				else
					Set RS1a = Server.CreateObject("ADODB.Recordset")
					SQL1a = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& materia &"' order by NU_Ordem_Boletim"
					RS1a.Open SQL1a, CON0
						
					if RS1a.EOF then
					else
					co_materia_fil_check=1 
					peso_acumula=0
					va_m31_acumula=0
					va_m32_acumula=0
					va_m33_acumula=0
					va_m34_acumula=0
					sem_nota1="n"
					sem_nota2="n"
					sem_nota3="n"
					sem_nota4="n"					
						while not RS1a.EOF
							co_mat_fil= RS1a("CO_Materia")
							
							Set RSp2 = Server.CreateObject("ADODB.Recordset")
							SQLp2 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia = '"& co_mat_fil &"' order by NU_Ordem_Boletim"
							RSp2.Open SQLp2, CON0	
													
							nu_peso_fil=RSp2("NU_Peso")	
										
							peso_acumula=peso_acumula+nu_peso_fil
							
							for periodofil=1 to 4
														
								Set RS3 = Server.CreateObject("ADODB.Recordset")
								SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& co_mat_fil &"' AND CO_Materia_Principal = '"& materia &"' AND NU_Periodo="&periodofil
								Set RS3 = CON_N.Execute(SQL_N)						
						

								if RS3.EOF then
									if periodofil=1 then
										f1="&nbsp;"
										va_m31_temp="&nbsp;"
									elseif periodofil=2 then
										f2="&nbsp;"
										va_m32_temp="&nbsp;"
									elseif periodofil=3 then
										f3="&nbsp;"
										va_m33_temp="&nbsp;"
									elseif periodofil=4 then
										f4="&nbsp;"
										va_m34_temp="&nbsp;"
									end if	
								else
									if periodofil=1 then
										f1=RS3("NU_Faltas")						
										va_m31_temp=RS3("VA_Media3")
									elseif periodofil=2 then
										f2=RS3("NU_Faltas")						
										va_m32_temp=RS3("VA_Media3")
									elseif periodofil=3 then
										f3=RS3("NU_Faltas")						
										va_m33_temp=RS3("VA_Media3")
									elseif periodofil=4 then
										f4=RS3("NU_Faltas")
										va_m34_temp=RS3("VA_Media3")
									end if
								end if
							next	
							if isnull(va_m31_temp) or va_m31_temp="&nbsp;"  or va_m31_temp="" then
								sem_nota1="s"
							else
								va_m31_acumula=va_m31_acumula+va_m31_temp								
							end if	
							
							if isnull(va_m32_temp) or va_m32_temp="&nbsp;" or va_m32_temp="" then
								sem_nota2="s"
							else
								va_m32_acumula=va_m32_acumula+va_m32_temp	
							end if
							
							if isnull(va_m33_temp) or va_m33_temp="&nbsp;" or va_m33_temp="" then
								sem_nota3="s"
							else
								va_m33_acumula=va_m33_acumula+va_m33_temp	
							end if
							
							if isnull(va_m34_temp) or va_m34_temp="&nbsp;" or va_m34_temp="" then
								sem_nota4="s"
							else
								va_m34_acumula=va_m34_acumula+va_m34_temp	
							end if
						RS1a.MOVENEXT
						wend
						if sem_nota1="s" then
							va_m31="&nbsp;"
						else	
							va_m31=va_m31_acumula/peso_acumula
							va_m31=va_m31*10
								decimo = va_m31 - Int(va_m31)
								If decimo >= 0.5 Then
									nota_arredondada = Int(va_m31) + 1
									va_m31=nota_arredondada
								else
									nota_arredondada = Int(va_m31)
									va_m31=nota_arredondada											
								End If
							va_m31=va_m31/10	
							va_m31 = formatNumber(va_m31,1)									
						end if	
						
						if sem_nota2="s" then
							va_m32="&nbsp;"
						else	
							va_m32=va_m32_acumula/peso_acumula
							va_m32=va_m32*10
								decimo = va_m32 - Int(va_m32)
								If decimo >= 0.5 Then
									nota_arredondada = Int(va_m32) + 1
									va_m32=nota_arredondada
								else
									nota_arredondada = Int(va_m32)
									va_m32=nota_arredondada											
								End If
							va_m32=va_m32/10	
							va_m32 = formatNumber(va_m32,1)									
						end if

						if sem_nota3="s" then
							va_m33="&nbsp;"
						else	
							va_m33=va_m33_acumula/peso_acumula
							va_m33=va_m33*10
								decimo = va_m33 - Int(va_m33)
								If decimo >= 0.5 Then
									nota_arredondada = Int(va_m33) + 1
									va_m33=nota_arredondada
								else
									nota_arredondada = Int(va_m33)
									va_m33=nota_arredondada											
								End If
							va_m33=va_m33/10	
							va_m33 = formatNumber(va_m33,1)								
						end if
						
						if sem_nota4="s" then
							va_m34="&nbsp;"
						else	
							va_m34=va_m34_acumula/peso_acumula
							va_m34=va_m34*10
								decimo = va_m34 - Int(va_m34)
								If decimo >= 0.5 Then
									nota_arredondada = Int(va_m34) + 1
									va_m34=nota_arredondada
								else
									nota_arredondada = Int(va_m34)
									va_m34=nota_arredondada											
								End If
							va_m34=va_m34/10	
							va_m34 = formatNumber(va_m34,1)									
						end if							
					end if									
				end if					
					if isnull(f1) or f1="&nbsp;"  or f1="" then
					soma_f1=0
					else
					soma_f1=f1
					end if

					if isnull(f2) or f2="&nbsp;"  or f2="" then
					soma_f2=0
					else
					soma_f2=f2
					end if
					
					if isnull(f3) or f3="&nbsp;"  or f3="" then
					soma_f3=0
					else
					soma_f3=f3
					end if					
					
					soma_f=soma_f1+soma_f2+soma_f3
					
					if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
					dividendo1=0
					divisor1=0
					else
					dividendo1=va_m31
					divisor1=1
					end if	
					
					if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
					dividendo2=0
					divisor2=0
					else
					dividendo2=va_m32
					divisor2=1
					end if
					
					if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
					dividendo3=0
					divisor3=0
					else
					dividendo3=va_m33
					divisor3=1
					end if
					
					if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
					nota_aux_m2_1="&nbsp;"
					dividendo4=0
					divisor4=0
					else
					nota_aux_m2_1=va_m34
					dividendo4=va_m34
					divisor4=1
					end if
					
					dividendo1=dividendo1*1	
					dividendo2=dividendo2*1	
					dividendo3=dividendo3*1		
					divisor1=divisor1*1
					divisor2=divisor2*1
					divisor3=divisor3*1									
					dividendo_ma=dividendo1+dividendo2+dividendo3
					divisor_ma=divisor1+divisor2+divisor3
					divisor_m3=divisor1+divisor2+divisor3+(divisor4*2)				
					'response.Write(dividendo_ma&"<<")
					
					if divisor_ma<3 then
					ma="&nbsp;"
					else
					ma=dividendo_ma
					end if
					
					if ma="&nbsp;" then
					else
					ma=ma*10
							decimo = ma - Int(ma)
							If decimo >= 0.5 Then
								nota_arredondada = Int(ma) + 1
								ma=nota_arredondada
							else
								nota_arredondada = Int(ma)
								ma=nota_arredondada											
							End If
					ma=ma/10	
						ma = formatNumber(ma,1)						
					end if

					
					if ma="&nbsp;" then
					dividendo_m2=0
					divisor_m2=0
					else
					dividendo_m2=ma+(dividendo4*2)
					divisor_m2=1
					end if
					
					if divisor_m2=0 then
					m2="&nbsp;"
					else
					'response.Write(mf&"="&dividendo_mf&"/"&divisor_mf)
					m2=dividendo_m2
					end if
					
					if m2="&nbsp;" then
					else
					m3=m2/divisor_m3
					m3=m3*10
						decimo = m3 - Int(m3)
							If decimo >= 0.5 Then
								nota_arredondada = Int(m3) + 1
								m3=nota_arredondada
							else
								nota_arredondada = Int(m3)
								m3=nota_arredondada					
							End If
					m3=m3/10		
						m3 = formatNumber(m3,1)								
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
			%>
                      <tr> 
                        <td width="339" class="<%response.Write(cor)%>"> 
                          <%response.Write(no_materia)
						  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
						  %>
                        </td>
                        <td width="60" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                            <%
							if showapr1="s" and showprova1="s" then																	
								if divisor1=1 then
									if va_m31>=ntazl then	
										response.Write("<font color="&cor_nota_prt&">"&formatnumber(va_m31,1)&"</font>")
									elseif va_m31>=ntvml then	
										response.Write("<font color="&cor_nota_azl&">"&formatnumber(va_m31,1)&"</font>")					
									else
										response.Write("<font color="&cor_nota_vml&">"&formatnumber(va_m31,1)&"</font>")	
									end if	
								else
									response.Write(va_m31)	
								end if	
							else
								response.Write("&nbsp;")									
							end if							
							%>
                          </div></td>
                        <td width="60" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                          <%
							if showapr2="s" and showprova2="s" then																	
								if divisor2=1 then
									if va_m32>=ntazl then	
										response.Write("<font color="&cor_nota_prt&">"&formatnumber(va_m32,1)&"</font>")
									elseif va_m32>=ntvml then	
										response.Write("<font color="&cor_nota_azl&">"&formatnumber(va_m32,1)&"</font>")					
									else
										response.Write("<font color="&cor_nota_vml&">"&formatnumber(va_m32,1)&"</font>")	
									end if	
								else
									response.Write(va_m32)	
								end if		
							else
								response.Write("&nbsp;")									
							end if	
							%>
                          </div></td>
                        <td width="60" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                          <%
							if showapr3="s" and showprova3="s" then																	
								if divisor3=1 then
									if va_m33>=ntazl then	
										response.Write("<font color="&cor_nota_prt&">"&formatnumber(va_m33,1)&"</font>")
									elseif va_m33>=ntvml then	
										response.Write("<font color="&cor_nota_azl&">"&formatnumber(va_m33,1)&"</font>")					
									else
										response.Write("<font color="&cor_nota_vml&">"&formatnumber(va_m33,1)&"</font>")	
									end if	
								else
									response.Write(va_m33)	
								end if
							else
								response.Write("&nbsp;")									
							end if	
							%>
                          </div></td>
                        <td width="60" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                            <%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" then
								response.Write(ma)
							else
								response.Write("&nbsp;")	
							end if
							%>
                          </div></td>
                        <td width="60" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                          <%
							if showapr4="s" and showprova4="s" then																	
								if divisor4=1 then
									if va_m34>=ntazl then	
										response.Write("<font color="&cor_nota_prt&">"&formatnumber(va_m34,1)&"</font>")
									elseif va_m34>=ntvml then	
										response.Write("<font color="&cor_nota_azl&">"&formatnumber(va_m34,1)&"</font>")					
									else
										response.Write("<font color="&cor_nota_vml&">"&formatnumber(va_m34,1)&"</font>")	
									end if	
								else
									response.Write(va_m34)	
								end if	
							else
								response.Write("&nbsp;")									
							end if	
							%>
                          </div></td>
                        <td width="60" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                            <%
							'if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" then
								response.Write(m2)
							else
								response.Write("&nbsp;")	
							end if
							%>
                          </div></td>
                        <td width="60" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                            <%
							'if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s"  and showapr4="s" and showprova4="s" then
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" then
								if m2<>"&nbsp;" then
									if m3>=ntazl then	
										response.Write("<font color="&cor_nota_prt&">"&formatnumber(m3,1)&"</font>")
									elseif m3>=ntvml then	
										response.Write("<font color="&cor_nota_azl&">"&formatnumber(m3,1)&"</font>")					
									else
										response.Write("<font color="&cor_nota_vml&">"&formatnumber(m3,1)&"</font>")	
									end if	
								else
									response.Write(m3)	
								end if	
							else
								response.Write("&nbsp;")	
							end if
							%>
                          </div></td>
                        <td width="1">&nbsp;</td>
                        <td width="60" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                            <%													
							response.Write(f1)	
							%>
                          </div></td>
                        <td width="60" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                            <%													
							response.Write(f2)	
							%>
                          </div></td>
                        <td width="60" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                            <%													
							response.Write(f3)	
							%>
                          </div></td>
                        <td width="60" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                            <%													
							response.Write(soma_f)	
							%>
                          </div></td>
                      </tr>
                      <%
			check=check+1
			RSprog.MOVENEXT
			wend
end if		
			%>
                    </table></td>
                </tr>
                <tr>
                  <td colspan="2" 
>&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="2" 
><table width="200" height="20" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td class="tb_tit"><div align="center"><a href="#" class="impressao" onClick="MM_openBrWindow('imprime.asp?obr=<%=obr%>','','menubar=yes,scrollbars=yes,resizable=yes,width=1000,height=450')">Vers�o 
                        para impress�o</a></div></td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
  </tr>
  <tr>
    <td width="1000" height="41"><div align="center"><img src="../../img/rodape.jpg" width="1000" height="41" /></div></td>
  </tr>
</table>
</form>
</body>
</html>
