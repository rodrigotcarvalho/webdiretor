<!--#include file="../../inc/caminhos.asp"-->
<!--#include file="../../inc/funcoes.asp"-->
<!--#include file="../../inc/funcoes2.asp"-->
<!--#include file="../../inc/funcoes6.asp"-->

<%
nivel=2
tp=session("tp")
nome = session("nome") 
co_user = session("co_user")
grafico=request.form("grafico")
periodo=request.form("periodo")
co_mat_fil= request.form("mat_prin")

if opt="1" then
'periodo_check=request.form("periodo")
cod_cons= Session("aluno_selecionado")
else
cod_cons= Session("aluno_selecionado")
'periodo_check=1
end if
cod_cons= Session("aluno_selecionado")

obr=cod_cons&"?"&periodo_check

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

	Set CONa = Server.CreateObject("ADODB.Connection") 
	ABRIRa = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONa.Open ABRIRa
		
	SQL2 = "select * from TB_Usuario where CO_Usuario = " & cod_cons 
	set RS2 = CON.Execute (SQL2)
	
nome_aluno= RS2("NO_Usuario")

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod_cons
		RS.Open SQL, CON1


ano_aluno = RS("NU_Ano")
rematricula = RS("DA_Rematricula")
situacao = RS("CO_Situacao")
encerramento= RS("DA_Encerramento")
unidade= RS("NU_Unidade")
curso= RS("CO_Curso")
co_etapa= RS("CO_Etapa")
turma= RS("CO_Turma")
cham= RS("NU_Chamada")

call GeraNomes("PORT",unidade,curso,co_etapa,CON0)
no_unidade = session("no_unidade")
no_curso = session("no_curso")
no_etapa = session("no_etapa")

if isnull(periodo) or  periodo="" then
periodo =1
end if

if isnull(co_mat_fil) or  co_mat_fil="" then
		Set RS5a = Server.CreateObject("ADODB.Recordset")
		SQL5a = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim"
		RS5a.Open SQL5a, CON0

co_mat_fil= RS5a("CO_Materia")

end if	

vetor_co_periodo=periodo
vetor_materia=co_mat_fil
m_cons="VA_Media3"

	



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

function submit()  
{
   var f=document.forms[0]; 
      f.submit(); 
}
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
</head>

<body onLoad="MM_preloadImages(<%response.Write(swapload)%>)">
<form action="grafico.asp" method="post">
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
      <td height="998"><table width="200" height="100%" border="0" cellpadding="0" cellspacing="0">
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
            <td height="832" colspan="2" valign="top"> <div align="left"> <p><img src="../../img/graficos.jpg" width="700" height="30"> 
                <input name="grafico" type="hidden" id="grafico" value="<%response.write(grafico)%>">
                </p></div>
                
              <table width="100%" border="0" align="center" cellspacing="0" class="tb_corpo" dwcopytype="CopyTableRow">
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
                          <input name="cod" type="hidden" value="<%=cod_cons%>">
                          <%response.Write(cod_cons)%>
                          </font></td>
                        <td width="6%" height="10"> <div align="right"><font class="style3"> 
                            Nome: </font></div></td>
                        <td width="66%" height="10"><font class="style1"> 
                          <%response.Write(nome_aluno)%>
                          <input name="nome2" type="hidden" class="textInput" id="nome2"  value="<%response.Write(nome_aluno)%>" size="75" maxlength="50">
                          &nbsp;</font></td>
                      </tr>
                    </table></td>
                  <td class="style3">                
				  <% 
grafico=grafico*1
if grafico=2 then %>
<div align="center">DISCIPLINA</div>
<%END if%>
</td>
                </tr>
                <tr> 
                  <td height="10" bgcolor="#FFFFFF">&nbsp;</td>
                  <td valign="top"><div align="center"><font class="style1"> 
                      <%				  
grafico=grafico*1
if grafico=2 then 

%>

 
                        <select name="mat_prin" class="borda" onChange="MM_callJS('submit()')">
                          <%				  
			Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND IN_MAE=TRUE order by NU_Ordem_Boletim "
		RS5.Open SQL5, CON0

	

while not RS5.EOF
co_mat_prin= RS5("CO_Materia")


		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Materia where CO_Materia ='"& co_mat_prin &"'"
		RS7.Open SQL7, CON0
		
		no_mat_prin= RS7("NO_Materia")
		
		IF co_mat_fil=co_mat_prin THEN
		%>
                          <option value="<%=co_mat_prin%>" SELECTED> 
                          <%response.Write(no_mat_prin)%>
                          </option>
                          <%
ELSE						  
%>
                          <option value="<%=co_mat_prin%>"> 
                          <%response.Write(no_mat_prin)%>
                          </option>
                          <%
END IF
RS5.MOVENEXT
WEND			  
END if
%>
                        </select>
                      </font></div></td>
                </tr>
                <tr> 
                  <td colspan="2"><table width="100%" border="0" cellspacing="0">
                      <tr class="style3"> 
                        <td width="34" height="10"> <div align="center"> 
                            <%					  
call GeraNomes("PORT",unidade,curso,co_etapa,CON0)
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
                        <td width="137"> <div align="center"> <font class="style1"> 
                                                            <select name="periodo" class="borda" id="select5" onChange="MM_callJS('submit()')">
                                  <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")
periodo=periodo*1
NU_Periodo=NU_Periodo*1
IF periodo=NU_Periodo then
vetor_nome_periodo=NO_Periodo
%>
                                  <option value="<%=NU_Periodo%>" selected> 
                                  <%response.Write(NO_Periodo)%>
                                  </option>
                                  <%
else								  
%>
                                  <option value="<%=NU_Periodo%>"> 
                                  <%response.Write(NO_Periodo)%>
                                  </option>
                                  <%
end if
RS4.MOVENEXT
WEND

%>
                                </select>
                            </font> </div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td bgcolor="#FFFFFF">&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
              </table>
		<% 
        grafico=grafico*1
        if grafico=1 then 
            tipo_grafico="media_geral"
            titulo_grafico="Comparar a M&eacute;dia de todas as Disciplinas do trimestre com a Turma e a Etapa (ano/s&eacute;rie)"
        else
            tipo_grafico="media_turma"
            titulo_grafico="Comparar a M&eacute;dia de uma Disciplina do trimestre com a Turma e a Etapa (ano/s&eacute;rie)"
        end if
			
		Set RSFIL = Server.CreateObject("ADODB.Recordset")
		SQLFIL = "SELECT * FROM TB_Da_Aula where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'"
		RSFIL.Open SQLFIL, CON2

	notaFIL=RSFIL("TP_Nota")

if notaFIL ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na

elseif notaFIL="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb

elseif notaFIL ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc
else
		response.Write("ERRO")
end if

		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Controle"
		RS4.Open SQL4, CON
	
co_apr1=RS4("CO_apr1")
co_apr2=RS4("CO_apr2")
co_apr3=RS4("CO_apr3")
co_apr4=RS4("CO_apr4")

co_prova1=RS4("CO_prova1")
co_prova2=RS4("CO_prova2")
co_prova3=RS4("CO_prova3")
co_prova4=RS4("CO_prova4")

dados_periodos=split(vetor_co_periodo,"#!#")
for y =0 to ubound(dados_periodos)
	periodo=dados_periodos(y)		
	periodo=periodo*1
	if periodo=1 then
		if co_apr1="D"then
			show="n"
		else 
			show="s"
		end if
		
		if co_prova1="D"then
			show="n"
		else 
			show="s"
		end if
	elseif periodo=2 then
		if co_apr2="D"then
			show="n"
		else 
			show="s"
		end if
		if co_prova2="D"then
			show="n"
		else 
			show="s"
		end if					
	elseif periodo=3 then	
		if co_apr3="D"then
			show="n"
		else 
			show="s"
		end if
		if co_prova3="D"then
			show="n"
		else 
			show="s"
		end if
	elseif periodo=4 then	
		if co_apr4="D"then
			show="n"
		else 
			show="s"
		end if
		if co_prova4="D"then
			show="n"
		else 
			show="s"
		end if
	end if
	if show="n" then
		vetor_aluno_quadro=""
		vetor_etapa_quadro=""
		vetor_turma_quadro=""
	else
		vetor_aluno=calcula_medias(unidade, curso, co_etapa, turma, periodo, cod_cons, vetor_materia, CAMINHOn, notaFIL, m_cons, tipo_grafico)
	
	
	'retirar o #$# do vetor pois ainda n�o terminei de mont�-lo
		media_aluno=split(vetor_aluno,"#$#")
	
		if y=0 then
			vetor_aluno_quadro=media_aluno(0)
		else	
			vetor_aluno_quadro=vetor_aluno_quadro&"#!#"&media_aluno(0)
		end if
		
				Set RSt0 = Server.CreateObject("ADODB.Recordset")
				SQLt0 = "SELECT * FROM TB_Aluno_Esta_Turma where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"'"
				RSt0.Open SQLt0, CONa
		
			co_matric_alunos_etapa_check=1
			while not RSt0.EOF
			co_matricula= RSt0("CO_Matricula")
			
				if co_matric_alunos_etapa_check=1 then
					co_matric_alunos_etapa=co_matricula
				else
					co_matric_alunos_etapa=co_matric_alunos_etapa&","&co_matricula
				end if
			co_matric_alunos_etapa_check=co_matric_alunos_etapa_check+1
			RSt0.MOVENEXT
			wend
		
		vetor_etapa=calcula_medias(unidade, curso, co_etapa, turma, periodo, co_matric_alunos_etapa, vetor_materia, CAMINHOn, notaFIL, m_cons, tipo_grafico)
	
	'retirar o #$# do vetor pois ainda n�o terminei de mont�-lo
		media_etapa=split(vetor_etapa,"#$#")
			
		if y=0 then
			vetor_etapa_quadro=media_etapa(0)
		else	
			vetor_etapa_quadro=vetor_etapa_quadro&"#!#"&media_etapa(0)
		end if
	
	
		
				Set RSt1 = Server.CreateObject("ADODB.Recordset")
				SQLt1 = "SELECT * FROM TB_Aluno_Esta_Turma where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'"
				RSt1.Open SQLt1, CONa
		
			co_matric_alunos_turma_check=1
			while not RSt1.EOF
			co_matricula= RSt1("CO_Matricula")
			
				if co_matric_alunos_turma_check=1 then
					co_matric_alunos_turma=co_matricula
				else
					co_matric_alunos_turma=co_matric_alunos_turma&","&co_matricula
				end if
			co_matric_alunos_turma_check=co_matric_alunos_turma_check+1
			RSt1.MOVENEXT
			wend
			
		vetor_turma=calcula_medias(unidade, curso, co_etapa, turma, periodo, co_matric_alunos_turma, vetor_materia, CAMINHOn, notaFIL, m_cons, tipo_grafico)
	
	'retirar o #$# do vetor pois ainda n�o terminei de mont�-lo
		media_turma=split(vetor_turma,"#$#")	
		if y=0 then
			vetor_turma_quadro=media_turma(0)
		else	
			vetor_turma_quadro=vetor_turma_quadro&"#!#"&media_turma(0)
		end if
	
	end if
	
next


	vetor_aluno_quadro="Aluno#!#"&vetor_aluno_quadro&"#$#"
	vetor_etapa_quadro="Etapa#!#"&vetor_etapa_quadro&"#$#"
	vetor_turma_quadro="Turma#!#"&vetor_turma_quadro&"#$#"



vetor_linha_quadro=vetor_aluno_quadro&vetor_etapa_quadro&vetor_turma_quadro

%>
              <table width="100%" border="0" align="center" cellspacing="0" class="tb_corpo">
                <tr>
                  <td width="653" class="tb_tit"><%response.Write(titulo_grafico)%></td>
                  <td width="113" class="tb_tit"
></td>
                </tr>
              <tr><td colspan="2">
<%

linhas=Split(vetor_linha_quadro,"#$#")

periodo_exibe= split(vetor_nome_periodo,"#!#")

largura_tabela=(100*ubound(periodo_exibe))+100+70
%>
<table width="<%response.Write(largura_tabela)%>" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="70">&nbsp;</td>
<%For j=0 to ubound(periodo_exibe)%>
    <th class="tb_tit" width="100"><%response.Write(periodo_exibe(j))%></th>
<%next%>    
  </tr>
<%For k=0 to ubound(linhas)%>
  <tr>
	<%
	colunas=Split(linhas(k),"#!#")
	For m=0 to ubound(colunas)
		if m=0 then
		%>   
			<th class="style3" width="70"><%response.Write(colunas(m))%></th> 
		<%else
		%>    

			<td><div align="center"><font class="style1"> <%response.Write(colunas(m))%></font></div></th>
		<%
		end if
	next%>      
  </tr>
<%
next

'vetor_linha_grafico=vetor_aluno&vetor_etapa&vetor_turma
grafico=Split(vetor_linha_quadro,"#$#")
'response.Write(vetor_linha_grafico)
session("faixas")=grafico(0)&"#$#"&grafico(1)&"#$#"&grafico(2)
session("faixas")=replace(session("faixas"),",",".")
session("categorias")=vetor_nome_periodo
'response.end()
%>  
</table> 
  </td></tr>                
                <tr>
                  <td colspan="2"><table width="100%" border="0" cellspacing="0">
                    <tr>
                      <td height="10" class="tb_corpo"><DIV align="center">
                        <iframe src ="iframe.asp" frameborder ="0" width="700" height="400" align="middle"> </iframe>
                      </DIV></td>
                    </tr>
                  </table></td>
                </tr>
              </table></td>
          </tr>
  </table></td>
  </tr> 
  <tr> 
    <td width="1000" height="41"><div align="center"><img src="../../img/rodape.gif" width="1000" height="41" /></div></td>
  </tr></table>
  </form>
</body>
</html>
