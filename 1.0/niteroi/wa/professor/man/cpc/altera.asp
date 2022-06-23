<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%
nivel=4

opt=request.QueryString("opt")

autoriza=Session("autoriza")
Session("autoriza")=autoriza
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
nvg = session("chave")
ano_letivo = session("ano_letivo") 
co_usr = session("co_user")
chave=nvg
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
nivel=4
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
if opt="pg" then
cod_cons= request.QueryString("c")
periodo= request.QueryString("p")
else
cod_cons= request.form("coordenador")
periodo= request.form("periodo")
end if


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CONG = Server.CreateObject("ADODB.Connection") 
		ABRIRG = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONG.Open ABRIRG		


		Set CONp = Server.CreateObject("ADODB.Connection") 
		ABRIRp = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONp.Open ABRIRp
		
		Set RSCO = Server.CreateObject("ADODB.Recordset")
		SQLCO = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& cod_cons
		RSCO.Open SQLCO, CON

nome_coor = RSCO("NO_Usuario")
cod_cons = RSCO("CO_Usuario")



 call navegacao (CON,chave,nivel)
navega=Session("caminho")

call linkFuncao(CON,sistema_local,modulo,setor,funcao,nivel)
link_funcao=session("link_funcao")



	%>
<html>
<head>
<title>Web Acad&ecirc;mico</title>
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
MM_reloadPage(true);
function submitfuncao()  
{
   var f=document.forms[0]; 
      f.submit(); 
	  
}  function checksubmit()
{
  if (document.inclusao.etapa.value == "999999")
  {    alert("Por favor, selecione uma etapa!")
    document.inclusao.etapa.focus()
    return false
  }         	     
  return true
}
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script> 
</head> 
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" background="../../../../img/fundo.gif" topmargin="0" marginwidth="0" marginheight="0">
<%

call cabecalho(nivel)

%>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">		  				  
                  <tr>                    
            <td height="10" class="tb_caminho">
              <%
	  response.Write(navega)

%>

	</td>
  </tr>
    <tr>                   
    <td height="10"> 
      <%	call mensagens(4,609,0,0) 

%>
    </td>
                </tr>

          <tr> 
            <td valign="top"> 
             
                
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo">

                  <tr> 
                    
            <td>                
              <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">
                <tr class="tb_tit"> 
                    
                  <td width="653" height="15" class="tb_tit">
                    Coordenador</td>
                  </tr>
                  <tr> 
                    <td> <table width="100%" border="0" cellspacing="0">
                        <tr> 
                          <td width="9%" height="30" class="tb_subtit"> <div align="right">C&oacute;digo: 
                              </div></td>
                          <td width="11%" height="30"><font class="form_dado_texto"> 
                            <input name="cod_cons" type="hidden" value="<%=cod_cons%>">
                            <%response.Write(cod_cons)%>
                            <input name="tp" type="hidden" id="tp" value="P">
                            <input name="acesso" type="hidden" id="acesso" value="2">
                            <input name="nome_prof" type="hidden" id="nome_prof" value="<% =nome_coor%>">

                            </font></td>
                          <td width="6%" height="30" class="tb_subtit"> <div align="right">Nome: 
                            </div></td>
                          <td width="74%" height="30"><font class="form_dado_texto"> 
                            <%response.Write(nome_coor)%>
                            </font> </td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr class="tb_tit"> 
                    
                  <td height="15" class="tb_tit"> Grade de Aulas</td>
                  </tr>
                  <tr> 
                    <td><table width="100%" border="0" cellspacing="0">
                      <tr> 
                        <td width="70" class="tb_subtit"> <div align="center">UNIDADE 
                          </div></td>
                        <td width="70" class="tb_subtit"> <div align="center">CURSO 
                          </div></td>
                        <td width="70" class="tb_subtit"> <div align="center">ETAPA 
                          </div></td>
                        <td width="70" class="tb_subtit"> <div align="center">TURMA 
                          </div></td>
                        <td width="260" class="tb_subtit"> 
                          <div align="center">PROFESSOR</div></td>
                        <td width="210" class="tb_subtit"> 
                        <div align="center">DISCIPLINA</div></td>
                        <td class="tb_subtit"><div align="center">BIM 1</div></td>
                        <td class="tb_subtit"><div align="center">BIM 2</div></td>
                        <td class="tb_subtit"><div align="center">BIM 3</div></td>
                        <td class="tb_subtit"><div align="center">BIM 4</div></td>
                        <td class="tb_subtit"><div align="center">FINAL</div></td>
                      </tr>
                      <%
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Da_Aula where CO_Cord ="& cod_cons &" order by CO_Curso, CO_Etapa, CO_Turma"
		RS1.Open SQL1, CONG, 3, 3
		

if RS1.EOF THEN
ELSE
check=2
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
cod_prof = RS1("CO_Professor")
curso = RS1("CO_Curso")
unidade = RS1("NU_Unidade")
co_etapa= RS1("CO_Etapa")
turma= RS1("CO_Turma")
mat_prin = RS1("CO_Materia_Principal")
mat_fil = RS1("CO_Materia")
tabela = RS1("TP_Nota")
'coordenador= RS1("CO_Cord")
		
		valor = unidade&"-"&curso&"-"&co_etapa&"-"&turma&"-"&mat_prin&"-"&mat_fil&"-"&tabela&"-"&coordenador
'response.Write("->"&cod_prof)

	Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RSu.Open SQLu, CON0
		
no_unidade = RSu("NO_Unidade")

		Set RSc = Server.CreateObject("ADODB.Recordset")
		SQLc = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RSc.Open SQLc, CON0
		
no_curso = RSc("NO_Abreviado_Curso")

	Set RSPR = Server.CreateObject("ADODB.Recordset")
		SQLPR = "SELECT * FROM TB_Professor where CO_Professor="& cod_prof 
		RSPR.Open SQLPR, CONp
		
if RSPR.eof then
professor=""
else		
professor = RSPR("NO_Professor")
end if

 if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if
%>
                      <tr> 
                        <td width="70" class="<%=cor%>"> <div align="center"> 
                            <font class="form_dado_texto"> 
                            <%response.Write(no_unidade)%>
                            </font></div></td>
                        <td width="70" class="<%=cor%>"> <div align="center"> 
                            <font class="form_dado_texto"> 
                            <%
response.Write(no_curso)%>
                            </font></div></td>
                        <td width="70" class="<%=cor%>"> <div align="center"> 
                            <font class="form_dado_texto"> 
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
                        <td width="70" class="<%=cor%>"> <div align="center"> 
                            <font class="form_dado_texto"> 
                            <%response.Write(turma)%>
                            </font></font></div></td>
                        <td width="260" class="<%=cor%>"> 
                          <div align="center"> 
                            <font class="form_dado_texto"> 
                            <%response.Write(professor)%>
                            </font></font></div></td>
                        <td width="210" class="<%=cor%>"> 
                          <div align="center"><font class="form_dado_texto"> 
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
                        </font></font> </div></td>
                        <td width="50" class="<%=cor%>"> <div align="center"> 
                            <%					

		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Da_Aula where CO_Professor="& cod_prof &"AND CO_Materia_Principal='"& mat_prin &"'AND NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		RS4.Open SQL4, CONG



if RS4.EOF then

		else	
p1 = RS4("ST_Per_1")
if p1 = "x" then
%>
                            <div align="center"><img src="../../../../img/s.gif" width="8" height="8" border="0"></div>
                            <%
else
%>
                            <div align="center"><img src="../../../../img/n.gif" width="8" height="8" border="0"></div>
                            <%
end if
end if
%>
                          </div></td>
                        <td width="50" class="<%=cor%>"> 
                          <%				

		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Da_Aula where CO_Professor="& cod_prof &"AND CO_Materia_Principal='"& mat_prin &"'AND NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		RS4.Open SQL4, CONG					
					
if RS4.EOF then

else	

p2 = RS4("ST_Per_2")
if p2 = "x" then
%>
                          <div align="center"><img src="../../../../img/s.gif" width="8" height="8" border="0"></div>
                          <%
else
%>
                          <div align="center"><img src="../../../../img/n.gif" width="8" height="8" border="0"></div>
                          <%
end if
end if

%>
                        </td>
                        <td width="50" class="<%=cor%>"> 
                          <%					

		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Da_Aula where CO_Professor="& cod_prof &"AND CO_Materia_Principal='"& mat_prin &"'AND NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		RS4.Open SQL4, CONG					
if RS4.EOF then
	
else	
p3 = RS4("ST_Per_3")
if p3 = "x" then
%>
                          <div align="center"><img src="../../../../img/s.gif" width="8" height="8" border="0"></div>
                          <%
else
%>
                          <div align="center"><img src="../../../../img/n.gif" width="8" height="8" border="0"></div>
                          <%
end if
end if
%>
                        </td>
                        <td width="50" class="<%=cor%>"> 
                          <%					

		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Da_Aula where CO_Professor="& cod_prof &"AND CO_Materia_Principal='"& mat_prin &"'AND NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		RS4.Open SQL4, CONG

if RS4.EOF then	

else	
p4 = RS4("ST_Per_4")
if p4 = "x" then
%>
                          <div align="center"><img src="../../../../img/s.gif" width="8" height="8" border="0"></div>
                          <%
else
%>
                          <div align="center"><img src="../../../../img/n.gif" width="8" height="8" border="0"></div>
                          <%
end if
end if

%>
                        </td>
                        <td width="50" class="<%=cor%>"><%					

		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Da_Aula where CO_Professor="& cod_prof &"AND CO_Materia_Principal='"& mat_prin &"'AND NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		RS4.Open SQL4, CONG

if RS4.EOF then	

else	
p4 = RS4("ST_Per_4")
if p4 = "x" then
%>
                          <div align="center"><img src="../../../../img/s.gif" alt="s" width="8" height="8" border="0"></div>
                          <%
else
%>
                          <div align="center"><img src="../../../../img/n.gif" alt="n" width="8" height="8" border="0"></div>
                          <%
end if
end if

%></td>
                      </tr>
                      <%
check=check+1
intrec = intrec + 1
RS1.MOVENEXT
WEND
END IF
END IF				
%>
                    </table></td>
                  </tr>
                </TABLE>
              </td>
                  </tr>
       <tr>
          <td ><div align="center">
		  <%for i=1 to RS1.PageCount
		  intpagina=intpagina*1
			  if i=intpagina then%>
			  <font class="form_dado_texto"><%response.Write(intpagina)%></font>
			  <%else%>
			   <a href="altera.asp?pagina=<%=response.Write(i)%>&nvg=<%=nvg%>&p=<%=periodo%>&c=<%=cod_cons%>&opt=pg" class="linkPaginacao"><%response.Write(i)%></a> 
			  <%end if
		  next
		  %></div></td>
        </tr>
                <tr> 
                  <td class="tb_tit"> <div align="center">
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
                </table>
              </td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
</body>
<%Call GravaLog (chave,cod_cons) %>
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