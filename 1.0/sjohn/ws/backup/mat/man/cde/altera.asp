<%'On Error Resume Next%>
<% Response.Charset="ISO-8859-1" %>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->


<%
opt= request.QueryString("opt")
ori= request.QueryString("ori")
pagina=Request.QueryString("pagina")
cod_cons= request.QueryString("cod_cons")

if cint(Request.QueryString("pagina"))<1 then
pagina=1
end if

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo")
ano_letivo_real = ano_letivo
sistema_local=session("sistema_local")

chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
	

nvg = session("chave")
chave=nvg
session("chave")=chave
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano_info=nivel&"-"&chave&"-"&ano_letivo

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1

		Set CON_al = Server.CreateObject("ADODB.Connection") 
		ABRIR_al = "DBQ="& CAMINHOa& ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_al.Open ABRIR_al

		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0


 call navegacao (CON,chave,nivel)
navega=Session("caminho")	

' Set RS2 = Server.CreateObject("ADODB.Recordset")
'		SQL2 = "SELECT * FROM TB_Autoriz_Usuario_Grupo Where CO_Usuario = "&cod_cons
'		RS2.Open SQL2, CON
		
'if RS2.EOF then

'else		
'co_grupo=RS2("CO_Grupo")
'End if

if ori="2" then


unidade_pesquisa=SESSION("unidade_pesquisa")
curso_pesquisa=SESSION("curso_pesquisa")
etapa_pesquisa=SESSION("etapa_pesquisa")
turma_pesquisa=SESSION("turma_pesquisa")	
doc_pesquisa=SESSION("doc_pesquisa")

	SESSION("unidade_pesquisa")=unidade_pesquisa
	SESSION("curso_pesquisa")=curso_pesquisa
	SESSION("etapa_pesquisa")=etapa_pesquisa
	SESSION("turma_pesquisa")=turma_pesquisa
	SESSION("doc_pesquisa")=doc_pesquisa

if unidade_pesquisa=999990 then
pontos_u=0
else
sql_unidade="NU_Unidade="&unidade_pesquisa&" AND "
pontos_u=5
end if

if curso_pesquisa="999990" then
pontos_c=0
else
sql_curso="CO_Curso='"&curso_pesquisa&"' AND "
pontos_c=10
end if	

if etapa_pesquisa="999990" then
pontos_e=0
else
sql_etapa="CO_Etapa='"&etapa_pesquisa&"' AND "
pontos_e=20
end if	

if  turma_pesquisa="999990" then
pontos_t=0
else
sql_turma="CO_Turma='"&turma_pesquisa&"' AND "
pontos_t=30
end if	

if doc_pesquisa="999990" then
pontos_d=0
else
sql_doc="where CO_Documento='"&doc_pesquisa&"'"
pontos_d=40
end if
pontuacao=pontos_u+pontos_c+pontos_e+pontos_t+pontos_d

					Set RSdt = Server.CreateObject("ADODB.Recordset")
					SQLdt = "SELECT * FROM TB_Documentos_Matricula "&sql_doc 
					RSdt.Open SQLdt, CON0
					
					while not RSdt.EOF
					doc=RSdt("CO_Documento")
					no_doc_mat=RSdt("NO_Documento")
					vetor_doc=vetor_doc&"##"&doc&"!!"&no_doc_mat				
					RSdt.MOVENEXT
					wend 

pesquisa_turmas= sql_unidade&sql_curso&sql_etapa&sql_turma

else
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod_cons
		RS.Open SQL, CON1
		
		
cod_cons = RS("CO_Matricula")
nome_aluno= RS("NO_Aluno")
end if
Call LimpaVetor2

%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="file:../../../../img/mm_menu.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
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
function submitforminterno()  
{
   var f=document.forms[3]; 
      f.submit(); 
	  
}
//-->
</script>
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
//-->
</script>
                         
<script>
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0" <%if (aluno_novo="s" and aluno_novo_dados="s") or aluno_novo="n" then%> onLoad="recuperarCursoLoad(<%response.Write(unidade_combo)%>);recuperarEtapaLoad(<%response.Write(curso_combo)%>);recuperarTurmaLoad(<%response.Write(etapa_altera)%>)" <%end if%>>
<% call cabecalho (nivel)
	  %>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
                    
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
	  </td>
	  </tr>
 <%if opt="ok" then%> 
              <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,401,2,0) %>
    </td>
  </tr>
 <%elseif opt="ok1" then%> 
              <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,9709,2,0) %>
    </td>
  </tr>  
 <%end if%> 
 <% if ori="2" then
 %>
             <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,420,0,0) 
	  %>
    </td>
  </tr>
 <%			
else
%>
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,302,0,0) 
	  %>
    </td>
  </tr>
 <%end if

 if ori="2" then

 %> 
<tr>

            <td valign="top"> 
			
<FORM name="formulario" METHOD="POST" ACTION="bd.asp?opt=i" onSubmit="return checksubmit()">
        <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
          <tr> 
            <td width="841" class="tb_tit"
>Documentos n&acirc;o entregues</td>
            <td width="151" class="tb_tit"
> </td>
            <td width="2" class="tb_tit"
></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td height="10" colspan="3"> <table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="100" class="tb_subtit"> 
                    <div align="center">Unidade </div></td>
                  <td width="200" class="tb_subtit"> 
                    <div align="center">Curso </div></td>
                  <td width="100" class="tb_subtit"> 
                    <div align="center">Etapa </div></td>
                  <td width="70" class="tb_subtit"> 
                    <div align="center">Turma </div></td>
                  <td width="70" class="tb_subtit"> 
                    <div align="center">Matr&iacute;cula</div></td>
                  <td width="210" height="10" class="tb_subtit"> 
                    <div align="center">Nome</div></td>
                  <td width="250" class="tb_subtit"> 
                    <div align="center">Documentos n&atilde;o Entregues</div></td>
                </tr>
                <%
			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT * FROM TB_Matriculas WHERE "&pesquisa_turmas&" NU_Ano="&ano_letivo&" order by CO_Matricula"
			RS1.Open SQL1, CON1, 3, 3
				

				
			if RS1.EOF then
				intpagina=1
				sem_link=1
						%>
								<tr> 
								  <td colspan="7" valign="top"> <div align="center"><font class="style1"> 
									  <%response.Write("Não existem documentos para os critérios informados!")%>
									  </font></div></td>
								</tr>
			<%else 

    if cint(Request.QueryString("pagina"))<=1 then
	intpagina = 1
    else
		if cint(Request.QueryString("pagina"))>RS1.PageCount then  
	    intpagina = RS1.PageCount
        else
    	intpagina = Request.QueryString("pagina")
		end if
    end if   
if pontuacao<70 then	
 RS1.PageSize = 10
else 
  RS1.PageSize = 30
 end if 
  
if Request.QueryString("pagina")="" then
      intpagina = 1
else
    if cint(Request.QueryString("pagina"))<1 then
	intpagina = 1
    else
		if cint(Request.QueryString("pagina"))>RS1.PageCount then  
	    intpagina = RS1.PageCount
        else
    	intpagina = Request.QueryString("pagina")
		end if
    end if   
 end if   
				
						 RS1.AbsolutePage = intpagina
						intrec = 0				
						sem_link=0		
						check=2											
					while intrec<RS1.PageSize and not RS1.EOF
					exibe="s"	
						cod_cons=RS1("CO_Matricula")				
						unidade=RS1("NU_Unidade")
						curso=RS1("CO_Curso")
						etapa=RS1("CO_Etapa")
						turma=RS1("CO_Turma")			
					
							Set RS2 = Server.CreateObject("ADODB.Recordset")
							SQL2 = "SELECT * FROM TB_Alunos WHERE CO_Matricula="&cod_cons
							RS2.Open SQL2, CON1
							
							nome = RS2("NO_Aluno")	
					
							Set RS0 = Server.CreateObject("ADODB.Recordset")
							SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="&unidade
							RS0.Open SQL0, CON0
							
							if RS0.eof then
							NO_Unidade = ""
							else
							NO_Unidade = RS0("NO_Unidade")
							end if
							
							Set RS0b = Server.CreateObject("ADODB.Recordset")
							SQL0b = "SELECT * FROM TB_Curso where CO_Curso='"&curso&"'"
							RS0b.Open SQL0b, CON0
				
							if RS0b.eof then
							NO_Curso = ""
							else
							NO_Curso = RS0b("NO_Curso")
							end if			
										
							Set RS0c = Server.CreateObject("ADODB.Recordset")
							SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&curso&"' AND CO_Etapa='"&etapa&"'"
							RS0c.Open SQL0c, CON0
							
							if RS0c.eof then
							NO_Etapa = ""
							else
							NO_Etapa = RS0c("NO_Etapa")
							end if			
									
						doc_desvetorizado = Split(vetor_doc, "##")					
						for i=1 to ubound(doc_desvetorizado)
							doc_individual=doc_desvetorizado(i)
							cod_doc_individual = Split(doc_individual, "!!")
							cod_doc=cod_doc_individual(0)
							nome_doc=cod_doc_individual(1)
							
						 if check mod 2 =0 then
						  cor = "tb_fundo_linha_par" 
						 else cor ="tb_fundo_linha_impar"
						 end if													
								
										Set RSde = Server.CreateObject("ADODB.Recordset")
										SQLde = "SELECT * FROM TB_Documentos_Entregues where CO_Documento='"&cod_doc&"' and CO_Matricula="&cod_cons
										RSde.Open SQLde, CON0
								
										if RSde.EOF then			
						%>
										<tr class="<%response.write(cor)%>"> 
										  <td width="100" height="10">
															  <div align="center"><font class="form_corpo"> 
											  <%response.Write(NO_Unidade)%></font>
											</div></td>
										  <td width="200"> 
											<div align="center"><font class="form_corpo"> 
											  <%response.Write(NO_Curso)%></font>
											</div></td>
										  <td width="100" height="10">
															  <div align="center"><font class="form_corpo"> 
											  <%response.Write(NO_Etapa)%></font>
											</div>				  
										  </td>
										  <td width="70" height="10"> 
											<div align="center"><font class="form_corpo"> 
											  <%response.Write(turma)%></font>
											</div>				  
										  </td>
										  <td width="70"> 
											<div align="center"><font class="form_corpo"> 
											  <%response.Write(cod_cons)%></font>
											</div></td>
										  <td width="210" height="10"> 
											<div align="center"><font class="form_corpo"> 
											  <a href="altera.asp?ori=3&opt=i&pagina=<%response.Write(pagina)%>&cod_cons=<%response.Write(cod_cons)%>" class="ativos"><%response.Write(nome)%></a>
											  </font></div></td>
										  <td width="250"> 
											<div align="center"> <font class="form_corpo"> 
										<%response.Write(nome_doc)%>
										   </font> </div></td>
										</tr>
									<%
										end if
									check=check+1	
								next
					intrec = intrec + 1				
					RS1.MOVENEXT
					wend 
			end if					
					%>	
             </table></td>
          </tr>
                <tr>
                  <td colspan="7">
				  	<table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td class="tb_tit"><div align="center"> 
                         <%if sem_link=0 then%>
                            &nbsp; 
                            <% if intpagina>1 then%>
                            <a href="altera.asp?ori=2&pagina=<%=intpagina-1%>" class="linktres">Anterior</a> 
                            <%end if
								for contapagina=1 to RS1.PageCount 
									pagina=pagina*1
										IF contapagina=pagina then
											response.Write("<font class=style3>"&contapagina&"</font>")
										else%>
											<a href="altera.asp?ori=2&pagina=<%=contapagina%>" class="linktres"><%response.Write(contapagina)%></a> 
										<%end if
								next
								if StrComp(intpagina,RS1.PageCount)<>0 then %>
                            <a href="altera.asp?ori=2&pagina=<%=intpagina + 1%>" class="linktres">Próximo</a> 
                            <%end if
							else%>
                            &nbsp; 
                            <%end if	
    RS1.close
    Set RS1 = Nothing
    %>
                          </div></td>
                      </tr>		  
          <tr> 
            <td valign="top" colspan="3"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="1000"><hr width="1000"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="25%"> <div align="center"> 
                            <input name="SUBMIT5" type=button class="borda_bot3" onClick="MM_goToURL('parent','index.asp?nvg=WS-MA-MA-CDE');return document.MM_returnValue" value="Voltar">
                          </div></td>
                        <td width="25%"> <div align="center"> </div></td>
                        <td width="25%"> <div align="center"> </div></td>
                        <td width="25%"> <div align="center"> </div></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td valign="top" colspan="3">&nbsp;</td>
          </tr>
        </table>
      </form>
	  
	  </td>
          </tr>
<%
else	
	%>  
<tr>

            <td valign="top"> 
<% if ori="3" then
url="bd.asp?ori=3&pagina="&pagina
else
url="bd.asp"
end if
%>			
<FORM name="formulario" METHOD="POST" ACTION="<%response.Write(url)%>" onSubmit="return checksubmit()">
        <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
          <tr> 
            <td width="841" class="tb_tit"
>Dados do Aluno</td>
            <td width="151" class="tb_tit"
> </td>
            <td width="2" class="tb_tit"
></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td height="10" colspan="3"> <table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="14%" height="10"><font class="form_dado_texto">Matr&iacute;cula</font></td>
                  <td width="2%"><div align="center">:</div></td>
                  <td width="22%" height="10"><font class="form_corpo"> 
                    <input name="cod_cons" type="hidden" id="cod_cons" value="<%=cod_cons%>">
                    <%response.Write(cod_cons)%>
                    </font></td>
                  <td width="15%" height="10"><font class="form_dado_texto">Nome</font></td>
                  <td width="1%"> <div align="center">:</div></td>
                  <td width="46%" height="10"><font class="form_corpo"> 
                    <%response.Write(nome_aluno)%>
                    </font></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td valign="top" colspan="3"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="1000" class="tb_tit">Documentos Entregues</td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="33%"><div align="center"><font class="form_dado_texto">Documento 
                            </font></div></td>
                        <td width="34%"><div align="center"><font class="form_dado_texto">Situa&ccedil;&atilde;o 
                            </font></div></td>
                        <td width="33%"><div align="center"><font class="form_dado_texto">Data 
                            </font></div></td>
                      </tr>
                      <%
		Set RSdt = Server.CreateObject("ADODB.Recordset")
		SQLdt = "SELECT * FROM TB_Documentos_Matricula order by NO_Documento"
		RSdt.Open SQLdt, CON0

while not RSdt.EOF
co_doc_mat=RSdt("CO_Documento")
no_doc_mat=RSdt("NO_Documento")


		Set RSde = Server.CreateObject("ADODB.Recordset")
		SQLde = "SELECT * FROM TB_Documentos_Entregues where CO_Documento='"&co_doc_mat&"' And CO_Matricula="&cod_cons
		RSde.Open SQLde, CON0

IF RSde.EOF then
%>
                      <tr> 
                        <td width="33%"><div align="center"><font class="form_corpo"> 
                            <%response.Write(no_doc_mat)%>
                            </font></div></td>
                        <td width="34%"><div align="center"> 
                            <table width="50%" border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr> 
                                <td width="8%"><input name="<%response.Write(co_doc_mat)%>" type="radio" value="S"></td>
                                <td width="38%"><font class="form_corpo"> Entregue 
                                  </font></td>
                                <td width="7%"><input type="radio" name="<%response.Write(co_doc_mat)%>" value="N" checked></td>
                                <td width="47%"><font class="form_corpo"> Pendente 
                                  </font></td>
                              </tr>
                            </table>
                          </div></td>
                        <td width="33%"><div align="center"><font class="form_corpo"> 
                            </font></div></td>
                      </tr>
                      <%else
data_ent=RSde("DA_Entrega_Documento")
%>
                      <tr> 
                        <td width="33%"><div align="center"><font class="form_corpo"> 
                            <%response.Write(no_doc_mat)%>
                            </font></div></td>
                        <td width="34%"><div align="center"> 
                            <table width="50%" border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr> 
                                <td width="8%"><input name="<%response.Write(co_doc_mat)%>" type="radio" value="S" checked></td>
                                <td width="38%"><font class="form_corpo"> Entregue 
                                  </font></td>
                                <td width="7%"><input type="radio" name="<%response.Write(co_doc_mat)%>" value="N"></td>
                                <td width="47%"><font class="form_corpo"> Pendente 
                                  </font></td>
                              </tr>
                            </table>
                          </div></td>
                        <td width="33%"><div align="center"><font class="form_corpo"> 
                            <%response.Write(data_ent)%>
                            </font></div></td>
                      </tr>
                      <%
end if
RSdt.Movenext
wend
%>
                    </table></td>
                </tr>
                <tr> 
                  <td><hr width="1000"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="25%"> <div align="center">
						<%if ori="3" then%>
                            <input name="SUBMIT5" type=button class="borda_bot3" onClick="MM_goToURL('parent','altera.asp?ori=2&pagina=<%response.write(pagina)%>');return document.MM_returnValue" value="Voltar">						
						<%else%>
                            <input name="SUBMIT5" type=button class="borda_bot3" onClick="MM_goToURL('parent','index.asp?nvg=WS-MA-MA-CDE');return document.MM_returnValue" value="Voltar">
                         <%end if%>
						  </div></td>
                        <td width="25%"> <div align="center"> </div></td>
                        <td width="25%"> <div align="center"> </div></td>
                        <td width="25%"> <div align="center"> 
                            <input name="SUBMIT" type=SUBMIT class="borda_bot2" value="Confirmar">
                          </div></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td valign="top" colspan="3">&nbsp;</td>
          </tr>
        </table>
      </form>
	  
	  </td>
          </tr>
<%end if%>
		  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.gif" width="1000" height="40"></td>
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