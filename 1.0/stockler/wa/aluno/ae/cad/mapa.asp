<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<% 
cod_cons= request.form("cod")	
periodo= request.form("periodo")
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
ori = request.QueryString("or")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
	
obr=cod_cons&"_"&co_mat_fil&"_"&periodo

if periodo=1 then
	m_cons="VA_Mc1"
elseif periodo=2 then
	m_cons="VA_Mc2"
elseif periodo=3 then
	m_cons="VA_Mc3"
elseif periodo=4 then
	m_cons="VA_Mfinal"
elseif periodo=5 then
	m_cons="VA_Media3"
elseif periodo=6 then
	m_cons="VA_Media3"
end if



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
		
		Set CONa = Server.CreateObject("ADODB.Connection") 
		ABRIRa = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONa.Open ABRIRa		

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod_cons
		RS.Open SQL, CON1
		
		nome_cons = RS("NO_Aluno")



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
} function createXMLHTTP()
            {
                        try
                        {
                                   ajax = new ActiveXObject("Microsoft.XMLHTTP");
                        }
                        catch(e)
                        {
                                   try
                                   {
                                               ajax = new ActiveXObject("Msxml2.XMLHTTP");
                                               alert(ajax);
                                   }
                                   catch(ex)
                                   {
                                               try
                                               {
                                                           ajax = new XMLHttpRequest();
                                               }
                                               catch(exc)
                                               {
                                                            alert("Esse browser não tem recursos para uso do Ajax");
                                                            ajax = null;
                                               }
                                   }
                                   return ajax;
                        }
           
           
               var arrSignatures = ["MSXML2.XMLHTTP.5.0", "MSXML2.XMLHTTP.4.0",
               "MSXML2.XMLHTTP.3.0", "MSXML2.XMLHTTP",
               "Microsoft.XMLHTTP"];
               for (var i=0; i < arrSignatures.length; i++) {
                                                                          try {
                                                                                                             var oRequest = new ActiveXObject(arrSignatures[i]);
                                                                                                             return oRequest;
                                                                          } catch (oError) {
                                                                          }
                                      }
           
                                      throw new Error("MSXML is not installed on your system.");
                        }                                
						
						

						 function recuperarPeriodo()
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=p", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_p= oHTTPRequest.responseText;
resultado_p = resultado_p.replace(/\+/g," ")
resultado_p = unescape(resultado_p)
document.all.divPeriodo.innerHTML = resultado_p
																	   
                                                           }
                                               }

                                               oHTTPRequest.send();
                                   }
								   

 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
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
      <%call mensagens(nivel,18,0,0) %>
    </td>
			  </tr>			  
<form name="form1" method="post" action="mapa.asp?ori=01">
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
                  <td width="19%" height="10"> <div align="right"><font class="form_dado_texto">Matr&iacute;cula: 
                      </font></div></td>
                  <td width="9%" height="10"> <font class="form_dado_texto"> 
                    <input name="cod" type="hidden" value="<%=cod_cons%>">
                    <%response.Write(cod_cons)%>
                    </font></td>
                  <td width="6%" height="10"> <div align="right"><font class="form_dado_texto">Nome: 
                      </font></div></td>
                  <td width="66%" height="10"> <font class="form_dado_texto"> 
                    <%response.Write(nome_cons)%>
                    <input name="nome" type="hidden" class="textInput" id="nome"  value="<%response.Write(nome_cons)%>" size="75" maxlength="50">
                    &nbsp; </font></td>
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
                  <td width="33" height="10"> <div align="center"> 
                      <%
call GeraNomes("PORT",unidade,curso,co_etapa,CON0)
no_unidades = session("no_unidades")
no_grau = session("no_grau")
no_serie = session("no_serie")
%>
                      Ano</div></td>
                  <td width="81" height="10"> <div align="center">Matr&iacute;cula</div></td>
                  <td width="75" height="10" class="tb_subtit"> <div align="center">Cancelamento</div></td>
                  <td width="86" height="10"> <div align="center"> Situa&ccedil;&atilde;o</div></td>
                  <td width="54" height="10"> <div align="center">Chamada</div></td>
                </tr>
                <tr class="tb_corpo"
> 
                  <td width="33" height="10"> <div align="center"><font class="form_dado_texto"> 
                      <%response.Write(ano_aluno)%>
                      </font> </div></td>
                  <td width="81" height="10"> <div align="center"><font class="form_dado_texto"> 
                      <%response.Write(rematricula)%>
                      </font></div></td>
                  <td width="75" height="10"> <div align="center"><font class="form_dado_texto"> 
                      <%response.Write(encerramento)%>
                      </font></div></td>
                  <td width="86" height="10"> <div align="center"><font class="form_dado_texto"> 
                      <%
					
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
		RSCONTST.Open SQLCONTST, CON0
							
				no_situacao = RSCONTST("TX_Descricao_Situacao")	
					response.Write(no_situacao)%>
                      </font></div></td>
                  <td width="54" height="10"> <div align="center"><font class="form_dado_texto"> 
                      <%response.Write(cham)%>
                      </font></div></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td colspan="5" bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="5"> <table width="100%" border="0" cellspacing="0">
                <tr class="tb_subtit"> 
                  <td width="190" height="10"> 
                    <div align="center"> 
                      Unidade</div></td>
                  <td width="240" height="10"> 
                    <div align="center"> Curso</div></td>
                  <td width="190" height="10"> 
                    <div align="center"> Etapa</div></td>
                  <td width="190" height="10"> 
                    <div align="center"> Turma </div></td>
                  <td width="190"> 
                    <div align="center">Per&iacute;odo</div></td>
                </tr>
                <tr> 
                  <td width="190" height="10" class="tb_corpo"
> 
                    <div align="center"><font class="form_dado_texto"> 
                      <%response.Write(no_unidades)%>
                      </font></div></td>
                  <td width="240" height="10" class="tb_corpo"
> 
                    <div align="center"><font class="form_dado_texto"> 
                      <%response.Write(no_grau)%>
                      </font></div></td>
                  <td width="190" height="10" class="tb_corpo"
> 
                    <div align="center"><font class="form_dado_texto"> 
                      <%response.Write(no_serie)%>
                      </font></div></td>
                  <td width="190" height="10" class="tb_corpo"
> 
                    <div align="center"><font class="form_dado_texto"> 
                      <%response.Write(turma)%>
                      </font></div></td>
                  <td width="190"> 
                    <div id="divPeriodo" align="center"> 
                      <select name="periodo" class="select_style" id="periodo" onChange="MM_callJS('submitfuncao()')">				  
                        <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")
periodo=periodo*1
NU_Periodo=NU_Periodo*1
if NU_Periodo=periodo then
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
WEND%>
                  </select></div></td>
                </tr>
                <tr> 
                  <td height="10" colspan="5" class="tb_corpo"
><hr></td>
                </tr>
                <tr>
                  <td height="10" colspan="5" class="tb_corpo"
>
                    <%
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
		RS5.Open SQL5, CON0
co_materia_check=1
while not RS5.EOF
	co_mat_fil= RS5("CO_Materia")				
	if co_materia_check=1 then
		'vetor_materia="'"&co_mat_fil&"'"
		vetor_materia=co_mat_fil
	else
		'vetor_materia=vetor_materia&",'"&co_mat_fil&"'"
		vetor_materia=vetor_materia&"#!#"&co_mat_fil		
	end if
	co_materia_check=co_materia_check+1			
			
RS5.MOVENEXT
wend	

vetor_materia_exibe=programa_aula(vetor_materia, unidade, curso, co_etapa, "nulo")	

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

'		Set CONn = Server.CreateObject("ADODB.Connection") 
'		ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
'		CONn.Open ABRIRn
'	
'		Set RSnFIL = Server.CreateObject("ADODB.Recordset")
'		SQLnFIL = "SELECT * FROM "&notaFIL&" where CO_Matricula ="& cod_cons &" AND CO_Materia in("& vetor_materia &") And NU_Periodo="&periodo
'		RSnFIL.Open SQLnFIL, CONn
'
'vetor_aluno="Aluno"
'while not RSnFIL.eof
''mat=RSnFIL("CO_Materia")
'media_aluno=RSnFIL("VA_Media3")
''response.Write(m_al&"_"&mat&"<br>")
'		if media_aluno="" or isnull(media_aluno) then
'		else
'		media_aluno=formatnumber(media_aluno,0)
'		end if 
'vetor_aluno=vetor_aluno&"#!#"&media_aluno
'RSnFIL.Movenext
'wend
'
vetor_quadro=calcula_medias(unidade, curso, co_etapa, turma, periodo, cod_cons, vetor_materia, CAMINHOn, notaFIL, m_cons, "media_turma")

vetor_aluno="Aluno#!#"&vetor_quadro

'response.end()
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
	
vetor_quadro=calcula_medias(unidade, curso, co_etapa, turma, periodo, co_matric_alunos_etapa, vetor_materia, CAMINHOn, notaFIL, m_cons, "media_turma")

vetor_etapa="Etapa#!#"&vetor_quadro
'	co_materia= split(vetor_materia,",")
'
'	vetor_etapa="Etapa"
'	for i=0 to ubound(co_materia)
'			Set RSnFIL2 = Server.CreateObject("ADODB.Recordset")
'			SQLnFIL2 = "SELECT Avg("&notaFIL&"."&m_cons&")AS MediaDeVA_Media3 FROM "&notaFIL&" where CO_Matricula in("&co_matric_alunos_etapa&") AND CO_Materia ="& co_materia(i)&" And NU_Periodo="&periodo
'			RSnFIL2.Open SQLnFIL2, CONn
'			
'		media_etapa=RSnFIL2("MediaDeVA_Media3")
'		if media_etapa="" or isnull(media_etapa) then
'		else
'		media_etapa=formatnumber(media_etapa,0)
'		end if 
'		vetor_etapa=vetor_etapa&"#!#"&media_etapa
'
'	NEXT
'
'
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
	
vetor_quadro=calcula_medias(unidade, curso, co_etapa, turma, periodo, co_matric_alunos_turma, vetor_materia, CAMINHOn, notaFIL, m_cons, "media_turma")

vetor_turma="Turma#!#"&vetor_quadro
'	vetor_turma="Turma"
'	for i=0 to ubound(co_materia)
'			Set RSnFIL2 = Server.CreateObject("ADODB.Recordset")
'			SQLnFIL2 = "SELECT Avg("&notaFIL&"."&m_cons&")AS MediaDeVA_Media3 FROM "&notaFIL&" where CO_Matricula in("& co_matric_alunos_turma &") AND CO_Materia ="& co_materia(i)&" And NU_Periodo="&periodo
'			RSnFIL2.Open SQLnFIL2, CONn
'			
'		media_turma=RSnFIL2("MediaDeVA_Media3")
'		if media_turma="" or isnull(media_turma) then
'		else
'		media_turma=formatnumber(media_turma,0)
'		end if 
'		vetor_turma=vetor_turma&"#!#"&media_turma
'	NEXT
	
vetor_linha_quadro=vetor_aluno&vetor_etapa&vetor_turma

linhas=Split(vetor_linha_quadro,"#$#")

co_materia_exibe= split(vetor_materia_exibe,"#!#")

largura_tabela=(50*ubound(co_materia_exibe))+50+70
%>
<table width="<%response.Write(largura_tabela)%>" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="70">&nbsp;</td>
<%For j=0 to ubound(co_materia_exibe)%>
    <th class="tb_tit" width="50"><%response.Write(co_materia_exibe(j))%></th>
<%next%>    
  </tr>
<%For k=0 to ubound(linhas)%>
  <tr>
	<%
	colunas=Split(linhas(k),"#!#")
	For m=0 to ubound(colunas)
		if m=0 then
		%>   
			<th class="tb_subtit" width="70"><%response.Write(colunas(m))%></th> 
		<%else
		%>    

			<td class="form_dado_texto"><div align="center"><%response.Write(colunas(m))%></div></th>
		<%
		end if
	next%>      
  </tr>
<%
next

session("faixas")=linhas(0)&"#$#"&linhas(1)&"#$#"&linhas(2)
session("categorias")=vetor_materia_exibe

%>  
</table>  
</td>
        </tr> 
        <tr> 
         <td colspan="5"> 
<DIV align="center">
<iframe src ="iframe.asp" frameborder ="0" width="1000" height="400" align="middle"> </iframe>
</DIV>         
         </td>
        </tr> 
              </table></td>
          </tr>
        </table>
      </td>
    </tr>
</form>
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

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
response.redirect("../../../../inc/erro.asp")
end if
%>