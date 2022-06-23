<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%
opt = request.QueryString("opt")

periodo=request.form("periodo")
nivel=4
nvg = session("nvg")
opt=request.QueryString("opt")
if opt="pg" then
cod_cons= request.QueryString("c")
periodo= request.QueryString("p")
else
cod_cons= request.form("coor")
periodo= request.form("periodo")
end if



ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("nvg")
session("nvg")=chave
session("chave")=session("nvg")


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2
				
 call navegacao (CON,chave,nivel)
navega=Session("caminho")	

		Set RSCO = Server.CreateObject("ADODB.Recordset")
		SQLCO = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& cod_cons
		RSCO.Open SQLCO, CON

nome_coor = RSCO("NO_Usuario")
cod = RSCO("CO_Usuario")


	%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
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
MM_reloadPage(true);//-->
</script>
</head> 
<body background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="alteracao" method="post" action="grade_cp2.asp">
<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
        <td height="10" class="tb_caminho"><font class="style-caminho"> 
              <%
	  response.Write(navega)

%>
      </font></td>
	  </tr>
          <tr> 
            
    <td height="10"> 
      <%	call mensagens(nivel,635,0,0) 
%>
    </td>
          </tr>
          <tr> 
            <td width="770" valign="top">
                
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
          <tr class="tb_tit"
> 
            <td width="653" height="15" class="tb_tit"
>Coordenador</td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="9%" height="30"> <div align="right">  <font class="form_dado_texto"><strong>C&oacute;digo: 
                      </strong></font></div></td>
                  <td width="11%" height="30">  <font class="form_dado_texto"> 
                    <%response.Write(cod_cons)%>
                    </font></td>
                  <td width="6%" height="30">   <font class="form_dado_texto"><div align="right"><strong>Nome: 
                   </strong> </div></font></td>
                  <td width="74%" height="30">  <font class="form_dado_texto"> 
                    <%response.Write(nome_coor)%>
                    </font></td>
                </tr>
              </table></td>
          </tr>
          <tr class="tb_tit"
> 
            <td height="15" class="tb_tit"
>Grade de Aulas</td>
          </tr>
          <tr> 
            <td><table width="1000" border="0" cellspacing="0">
                <tr> 
                  <td width="72" class="tb_subtit"> <div align="center">UNIDADE 
                    </div></td>
                  <td width="72" class="tb_subtit"> <div align="center">CURSO 
                    </div></td>
                  <td width="72" class="tb_subtit"> <div align="center">ETAPA 
                    </div></td>
                  <td width="72" class="tb_subtit"> <div align="center">TURMA 
                    </div></td>
                  <td width="312" class="tb_subtit"> <div align="center">PROFESSOR</div></td>
                  <td width="250" class="tb_subtit"> <div align="center">DISCIPLINA</div></td>
                  <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo where NU_Periodo="&periodo
		RS4.Open SQL4, CON0

NO_Periodo= RS4("NO_Periodo") 
						  
						  %>
                  <td width="150" class="tb_subtit"> <div align="center"> 
                      <%response.Write(ucase(NO_Periodo))%>
                    </div></td>
                </tr>
                <%
if periodo = "1" then
periodo_SQL="ST_Per_1"
elseif periodo = "2" then
periodo_SQL="ST_Per_2"
elseif periodo = "3" then
periodo_SQL="ST_Per_3"
elseif periodo = "4" then
periodo_SQL="ST_Per_4"
elseif periodo = "5" then
periodo_SQL="ST_Per_5"
elseif periodo = "6" then
periodo_SQL="ST_Per_6"
end if

				
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Da_Aula where CO_Cord ="& cod_cons &" AND (isnull("&periodo_SQL&") or "&periodo_SQL&"='') order by CO_Curso, CO_Etapa, CO_Turma"
		RS1.Open SQL1, CON1, 3, 3



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
 
 RS1.PageSize = 28
 
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
		

if RS1.EOF THEN
intpagina=1
sem_link=1
ELSE


check=2

    RS1.AbsolutePage = intpagina
    intrec = 0
	
While intrec<RS1.PageSize and not RS1.EOF
codigo = RS1("CO_Professor")
curso = RS1("CO_Curso")
unidade = RS1("NU_Unidade")
co_etapa= RS1("CO_Etapa")
turma= RS1("CO_Turma")
mat_prin = RS1("CO_Materia_Principal")
mat_fil = RS1("CO_Materia")
tabela = RS1("TP_Nota")

 if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if
  
		valor = unidade&"-"&curso&"-"&co_etapa&"-"&turma&"-"&mat_prin&"-"&mat_fil&"-"&tabela&"-"&coordenador


	Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RSu.Open SQLu, CON0
		
no_unidade = RSu("NO_Unidade")

		Set RSc = Server.CreateObject("ADODB.Recordset")
		SQLc = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RSc.Open SQLc, CON0
		
no_curso = RSc("NO_Abreviado_Curso")

	Set RSPR = Server.CreateObject("ADODB.Recordset")
		SQLPR = "SELECT * FROM TB_Professor where CO_Professor="& codigo 
		RSPR.Open SQLPR, CON2
		
if RSPR.eof then
professor=""
else		
professor = RSPR("NO_Professor")
end if		

%>
                <tr> 
                  <td width="72" class="<%=cor%>"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_unidade)%>
                      </font></div></td>
                  <td width="72" class="<%=cor%>"> <div align="center"> <font class="form_dado_texto"> 
                      <%
response.Write(no_curso)%>
                      </font></div></td>
                  <td width="72" class="<%=cor%>"> <div align="center"> <font class="form_dado_texto"> 
                      <%

		Set RSe = Server.CreateObject("ADODB.Recordset")
		SQLe = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' and CO_Curso ='"& curso &"'" 
		RSe.Open SQLe, CON0
		
if RSe.EOF THEN
no_etapa="sem etapa"
else
no_etapa=RSe("NO_Etapa")
end if
response.Write(no_etapa)%>
                      </font></div></td>
                  <td width="72" class="<%=cor%>"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(turma)%>
                      </font></div></td>
                  <td width="312" class="<%=cor%>"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(professor)%>
                      </font></div></td>
                  <td width="250" class="<%=cor%>"> <div align="center"> <font class="form_dado_texto"> 
                      <%
		Set RSm = Server.CreateObject("ADODB.Recordset")
		SQLm = "SELECT * FROM TB_Materia where CO_Materia ='"& mat_prin &"'" 
		RSm.Open SQLm, CON0
		
if RSm.EOF THEN
no_mat_prin="sem disciplina"
else
no_mat_prin=RSm("NO_Materia")
end if
response.Write(no_mat_prin)%>
                      </font></div></td>
                  <td width="150" class="<%=cor%>"> <div align="center"> <font class="form_dado_texto"> 
                      <%					

%>
                      </font><font class="form_dado_texto"> 
                      <div align="center"><a href="../bdln/confirma.asp?or=02&cfp=cnpc&coor=<%=cod_cons%>&opt=blq&nt=<%=tabela%>&a=<%=ano_letivo%>&u=<% = unidade %>&c=<% = curso %>&e=<% = co_etapa %>&t=<% = turma%>&d=<%=mat_prin%>&pr=<%=codigo%>&p=1"><img src="../../../../img/n.gif" width="8" height="8" border="0"></a></div>
                      </font></div></td>
                  <%
check=check+1
intrec = intrec + 1
RS1.MOVENEXT
		
%>
                </tr>
                <%
WEND
END IF				
%>        <tr>
          <td colspan="7" ><div align="center">
		  <%for i=1 to RS1.PageCount
		  intpagina=intpagina*1
			  if i=intpagina then%>
			  <font class="form_dado_texto"><%response.Write(intpagina)%></font>
			  <%else%>
			   <a href="altera.asp?pagina=<%=response.Write(i)%>&nvg=<%=nvg%>&p=<%=periodo%>&c=<%=cod_cons%>&opt=pg" class="linkPaginacao"><%response.Write(i)%></a> 
			  <%end if
		  next
		  %></div>
		  </td>
        </tr>
                <tr> 
                  <td colspan="7" class="tb_tit"> <div align="center">
                    <%
if sem_link=0 then
		  %>&nbsp; <%
			    if intpagina>1 then
    %>
                      <a href="altera.asp?pagina=<%=intpagina-1%>&nvg=<%=nvg%>&p=<%=periodo%>&c=<%=cod_cons%>&opt=pg" class="linktres">Anterior</a> 
                      <%
    end if
    if StrComp(intpagina,RS1.PageCount)<>0 then  
    %>
                      <a href="altera.asp?pagina=<%=intpagina + 1%>&nvg=<%=nvg%>&p=<%=periodo%>&c=<%=cod_cons%>&opt=pg" class="linktres">Próximo</a> 
                      <%
    end if
else	
	%>
                      &nbsp; 
                      <%
end if	
    RS1.close
    Set RS1 = Nothing
    %>
                    </div></td>
                </tr>
              </table></td>
          </tr>
        </table>
              </td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
</form>
</body>
<%Call GravaLog (chave,"0") %>
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