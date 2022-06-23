<!--#include file="inc/caminhos.asp"-->
<!--#include file="inc/funcoes.asp"-->
<!--#include file="inc/funcoes2.asp"-->
<%
tp=session("tp")
nome = session("nome") 
co_user = session("co_user")
ano_letivo=session("ano_letivo")
nivel=0

 	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CON_wf = Server.CreateObject("ADODB.Connection")
 	ABRIR = "DBQ="& CAMINHO_wf& ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_wf.Open ABRIR
	
	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1
	
		Set CON3 = Server.CreateObject("ADODB.Connection") 
		ABRIR3 = "DBQ="& CAMINHO_o & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3	
	
		Set CON7 = Server.CreateObject("ADODB.Connection") 
		ABRIR7 = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON7.Open ABRIR7

opt=request.QueryString("opt")

			if tp="R" then			
			tipo="Responsável"
				SQL = "select * from TB_RespxAluno where CO_Usuario = " & co_user &" ORDER BY CO_Aluno"
				set RS = CON_wf.Execute (SQL)
				'quantos=RS("quantos")
				quantos=0
				alunos_vetor=0
				While not RS.EOF
					alunos=RS("CO_Aluno")
					if quantos=0 then
					Session("aluno_selecionado")=alunos
					end if
					alunos_vetor=alunos_vetor&"?"&alunos
					quantos=quantos+1
				RS.MOVENEXT
				WEND				
				if opt="as" then
				alunos=request.form("co_aluno")
				Session("aluno_selecionado")=alunos
				end if
				if opt="ad" or opt="sa" then
				alunos=Session("aluno_selecionado")
				Session("aluno_selecionado")=alunos
				end if		
			elseif tp="A" then
			tipo="Aluno"
			Session("aluno_selecionado")=co_user
			response.Redirect("aluno.asp")			
			end if
			
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Web Família</title>
<link href="estilo.css" rel="stylesheet" type="text/css" />
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
<table width="1000" height="1039" border="0" align="center" cellpadding="0" cellspacing="0" class="tabela_principal">
  <%
			ano = DatePart("yyyy", now)
			mes = DatePart("m", now) 
			dia = DatePart("d", now) 
			hora = DatePart("h", now) 
			min = DatePart("n", now) 
			semana = WeekDay(now)
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
          <td height="90" colspan="3"> 
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
          <td height="5" colspan="2"><p><img src="img/linha-pontilhada_grande.gif" alt="" width="828" height="5" /></p></td>
        </tr>
      <tr class="tabela_menu">
        <td height="19" colspan="2">&nbsp;</td>
      </tr>		
        <tr class="tabela_menu"> 
          <td height="832" colspan="2" valign="top"> <p align="left"><img src="img/inicial.jpg" width="700" height="30"></p>
            <form name="form1" method="post" action="inicio.asp?opt=as">
              <table width="800" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="10">&nbsp;</td>
                  <td height="30" colspan="7" valign="top"><font class="style3">O 
                    Web Fam&iacute;lia disponibilizar&aacute; acesso &agrave;s 
                    informa&ccedil;&otilde;es do aluno que estiver selecionado 
                    abaixo:</font></td>
                </tr>
                <tr> 
                  <td width="10">&nbsp;</td>
                  <td width="20">&nbsp;</td>
                  <td width="80"> <div align="center"><font class="style3">MATR&Iacute;CULA</font></div></td>
                  <td width="245"><font class="style3">NOME</font></td>
                  <td width="80"> <div align="center"><font class="style3">UNIDADE</font></div></td>
                  <td width="195"><font class="style3">CURSO</font></td>
                  <td width="60"> <div align="center"><font class="style3">TURMA</font></div></td>
                  <td width="109"><div align="center"><font class="style3">NASCIMENTO</font></div></td>
                </tr>
                <%
 	

vetor = split(alunos_vetor,"?")
for i =1 to ubound(vetor)
co_aluno= vetor(i)

	SQL2 = "select * from TB_Alunos where CO_Matricula = " & co_aluno 
	set RS2 = CON1.Execute (SQL2)
	
nome_aluno= RS2("NO_Aluno")

		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Contatos where CO_Matricula = " & co_aluno &" AND TP_Contato='ALUNO'"
		RS7.Open SQL7, CON7

nascimento = RS7("DA_Nascimento_Contato")

dados_dtd= split(nascimento, "/" )
dia_de= dados_dtd(0)
mes_de= dados_dtd(1)
ano_de= dados_dtd(2)
dia_de=dia_de*1
if dia_de<10 then
dia_de="0"&dia_de
end if
mes_de=mes_de*1
if mes_de<10 then
mes_de="0"&mes_de
end if

nascimento=dia_de&"/"&mes_de&"/"&ano_de

	SQL3 = "select * from TB_Matriculas where NU_Ano="& ano_letivo &" AND CO_Matricula = " & co_aluno 
	set RS3 = CON1.Execute (SQL3)

nu_unidade= RS3("NU_Unidade")
co_curso= RS3("CO_Curso")
co_etapa= RS3("CO_Etapa")
co_turma= RS3("CO_Turma")


'call GeraNomes("PORT",nu_unidade,co_curso,co_etapa,CON0)

    
no_unidade = GeraNomesNovaVersao("U",nu_unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)
no_curso = GeraNomesNovaVersao("C",co_curso,variavel2,variavel3,variavel4,variavel5,CON0,outro)
no_etapa = GeraNomesNovaVersao("E",co_curso,co_etapa,variavel3,variavel4,variavel5,CON0,outro)
prep_curso=GeraNomesNovaVersao("PC",co_curso,variavel2,variavel3,variavel4,variavel5,CON0,outro)
local= no_etapa&" "&prep_curso&" "&no_curso
%>
                <tr> 
                  <td width="10">&nbsp;</td>
                  <td> 
                    <%
					if opt="as" or opt="ad" or opt="sa"  then
alunos=Session("aluno_selecionado")
co_aluno=co_aluno*1
alunos=alunos*1
					if co_aluno=alunos then
						unidade_documentos=nu_unidade
						curso_documentos=co_curso
						etapa_documentos=co_etapa
						turma_documentos=co_turma					
					%>
                    <input name="co_aluno" type="radio" onClick="MM_callJS('submit()')" value="<%=co_aluno%>" checked> 
                    <%else%>
                    <input type="radio" name="co_aluno" onClick="MM_callJS('submit()')" value="<%=co_aluno%>">	
                    <%end if
else
					if i=1 then
					Session("aluno_selecionado")=co_aluno
					%>
                    <input name="co_aluno" type="radio" onClick="MM_callJS('submit()')" value="<%=co_aluno%>" checked> 
                    <%else%>
                    <input type="radio" name="co_aluno" onClick="MM_callJS('submit()')" value="<%=co_aluno%>">	
                    <%end if%>
                    <%end if%>
                  </td>
                  <td><div align="center"><font class="style1"> 
                      <%response.write(co_aluno)%>
                      </font> </div></td>
                  <td> <font class="style1"> 
                    <%response.write(nome_aluno)%>
                    </font> </td>
                  <td><div align="center"><font class="style1"> 
                      <%response.write(no_unidade)%>
                      </font></div></td>
                  <td><font class="style1"> 
                    <%response.write(local)%>
                    </font></td>
                  <td><div align="center"><font class="style1"> 
                      <%response.write(co_turma)%>
                      </font></div></td>
                  <td><div align="center"><font class="style1"> 
                      <%response.write(nascimento)%>
                      </font></div></td>
                </tr>

                <%
NEXT
%>
                <tr>
                  <td></td>
                  <td height="40" colspan="7" valign="bottom">
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <th colspan="2" scope="row"><hr>                      </th>
                      </tr>
<%
	Set RS4 = Server.CreateObject("ADODB.Recordset")
	SQL4 = "select * from TB_Ocorrencia_Aluno where CO_Matricula = " & alunos &" Order BY DA_Ocorrencia DESC,HO_Ocorrencia"
	set RS4 = CON3.Execute (SQL4)
if RS4.EOF then
%>
                    <tr>
                     <th width="190" scope="row"><div align="right"><font class="style3"> Última Ocorrencia Registrada:</font></div></th>
                      <td width="610" scope="value"><div align="left"><font class="style1">&nbsp;Sem Ocorrências</font></div></th>
                    </tr>
<%else
data_ocor=RS4("DA_Ocorrencia")
%>
                    <tr>
                      <th width="190" scope="row"><div align="right"><font class="style3"> Última Ocorrencia Registrada:</font></div></th>
                      <td width="610" scope="value"><div align="left"><font class="style1"> &nbsp;<%response.Write(data_ocor)%></font></div></td>
                    </tr>
<%end if
		Set RS_doc = Server.CreateObject("ADODB.Recordset")
		SQL_doc = "SELECT * FROM TB_Documentos where TP_Doc= 1 AND (((Unidade='"&unidade_documentos&"') AND (Curso='"&curso_documentos&"') AND  (Etapa='"&etapa_documentos&"') AND (Turma='"&turma_documentos&"')) OR ((Unidade='"&unidade_documentos&"') AND (Curso Is Null) AND (Etapa Is Null) AND (Turma Is Null)) OR ((Unidade='"&unidade_documentos&"') AND (Curso='"&curso_documentos&"') AND (Etapa Is Null) AND (Turma Is Null)) OR ((Unidade='"&unidade_documentos&"') AND (Curso='"&curso_documentos&"') AND  (Etapa='"&etapa_documentos&"') AND (Turma Is Null)) OR  ((Unidade Is Null) AND (Curso Is Null) AND  (Etapa Is Null) AND (Turma Is Null))) order by DA_Doc Desc"
		RS_doc.Open SQL_doc, CON_wf
if RS_doc.eof then
%>
                      <tr class="<%response.write(cor)%>"> 
                      <th width="190" scope="row"><div align="right"><font class="style3">Último Informe Escolar:</font></div> </th>
                        <td width="610" scope="value"> <div align="Left"><font class="style1"> 
                          &nbsp;Sem Publicações.
                          </font></div></td>
                      </tr>

<%else
tipo_arquivo=RS_doc("TP_Doc")
tit1=RS_doc("TI1_Doc")
data_pub=RS_doc("DA_Doc")

select case tipo_arquivo
case 1
nome_tipo_arquivo="Circulares"
case 2
nome_tipo_arquivo="Avaliações e Gabaritos"
case 3
nome_tipo_arquivo="Reunião de Pais"
end select

if data_pub="" or isnull(data_pub) then
else			
dados_dtd= split(data_pub, "/" )
dia_de= dados_dtd(0)
mes_de= dados_dtd(1)
ano_de= dados_dtd(2)
end if


if dia_de<10 then
dia_de="0"&dia_de
end if
if mes_de<10 then
mes_de="0"&mes_de
end if

data_inicio=dia_de&"/"&mes_de&"/"&ano_de
%>                    
                    <tr>
                      <th width="190" scope="row"><div align="right"><font class="style3">Último Informe Escolar:</font></div> </th>
                      <td width="610" scope="value"><font class="style1"> &nbsp;<%response.Write(nome_tipo_arquivo &" - "&tit1 &", publicado em "&data_inicio)%></font></td>
                    </tr>
<%end if%>                    
                  </table></td>
                </tr>
                <tr>
				  <td width="10"></td>
                  <td height="40" colspan="7" valign="bottom"> 
                    <%if quantos>1 then%>
                    <font class="style3"> ATEN&Ccedil;&Atilde;O ! </font><font class="style3">Caso 
                    queira obter informa&ccedil;&otilde;es de outro aluno volte 
                    a P&Aacute;GINA INICIAL e fa&ccedil;a nova sele&ccedil;&atilde;o.</font> 
                    <%end if%>
                  </td>
                </tr>
              </table>
            </form>            
          </td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td width="1000" height="40"><img src="img/rodape.jpg" width="1000" height="40" /></td>
  </tr>
</table>
</body>
</html>
