<%On Error Resume Next%>

<!--#include file="../../../../inc/funcoes.asp"-->



<!--#include file="../../../../inc/caminhos.asp"-->


<% 
opt=request.QueryString("opt")
ano_letivo = session("ano_letivo")
cod_prof = session("cod_prof")
nome_prof = session("nome_prof")
co_usr_prof = session("co_usr_prof")
grade = session("pendentes")
session("cod_prof")=cod_prof
session("nome_prof")=nome_prof
session("co_usr_prof")=co_usr_prof



		Set CON_g = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_g.Open ABRIR
					

nivel=4
nvg = session("chave")
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
chave=nvg
session("chave")=chave







		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0

 call navegacao (CON,chave,nivel)
navega=Session("caminho")


vertorExclui = split(grade,", ")
'RESPONSE.Write("<BR>->"&ubound(vertorExclui))
if ubound(vertorExclui)=-1 then
response.Redirect("grade_cp1.asp?opt=ok2&cod_cons="&cod_prof&"")
else

if opt="sw" then

for i =1 to ubound(vertorExclui)
'RESPONSE.Write("<BR>->"&vertorExclui(i))
exclui = split(vertorExclui(i),"-")
unidade = exclui(0)
curso= exclui(1)
co_etapa= exclui(2)
turma= exclui(3)
mat_prin= exclui(4)
mat_fil= exclui(5)
tabela= exclui(6)
coordenador= exclui(7)


if i=1 then
session("ano_letivo") = ano_letivo
session("cod_prof") = cod_prof
session("nome_prof") = nome_prof
session("co_usr_prof") = co_usr_prof
pendentes2 = grade


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
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</Head>
<body background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
    <td height="10" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)
	  %>
      </font></td>
  </tr>
  <tr> 
    <td  height="10" valign="top"> 
      <%
	call mensagens(nivel,634,0,0) 
%>
    </td>
  </tr>
  <tr> 
    <td valign="top"> <form action="bd2.asp?opt=de" method="post" name="alteracao">
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
          <tr class="tb_tit"
> 
            <td width="653" height="15" class="tb_tit"
>Professor</td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="9%" height="30"> <div align="right"><font class="tb_subtit">C&oacute;digo: 
                      </font></div></td>
                  <td width="11%" height="30"><font class="form_dado_texto"> 
                    <input name="cod" type="hidden" value="<%=cod_prof%>">
                    <%response.Write(cod_prof)%>
                    <input name="tp" type="hidden" id="tp" value="P">
                    <input name="acesso" type="hidden" id="acesso" value="2">
                    <input name="co_usr_prof" type="hidden" id="co_usr_prof" value="<% =co_usr_prof%>">
                    </font></td>
                  <td width="6%" height="30"> <div align="right"><font class="tb_subtit">Nome: 
                      </font></div></td>
                  <td width="74%" height="30"><font class="form_dado_texto"> 
                    <%response.Write(nome_prof)%>
                    </font></div> </td>
                </tr>
              </table></td>
          </tr>
          <tr class="tb_tit"
> 
            <td height="15" class="tb_tit"
>Grade de aulas</td>
          </tr>
          <tr> 
            <td><table width="1000" border="0" cellspacing="0">
                <tr> 
                  <td width="20">&nbsp;</td>
                  <td width="60" class="tb_subtit"> <div align="center"><strong>UNIDADE 
                      </strong></div></td>
                  <td width="45" class="tb_subtit"> <div align="center"><strong>CURSO 
                      </strong></div></td>
                  <td width="175" class="tb_subtit"> <div align="center"><strong>ETAPA 
                      </strong></div></td>
                  <td width="40" class="tb_subtit"> <div align="center"><strong>TURMA 
                      </strong></div></td>
                  <td width="190" class="tb_subtit"> <div align="center"><strong>DISCIPLINA</strong></div></td>
                  <td width="68" class="tb_subtit"> <div align="center"><strong>MODELO</strong></div></td>
                  <td width="171" class="tb_subtit"> <div align="center"><strong>COORDENADOR 
                      </strong></div></td>
                </tr>
                <%

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Abreviado_Curso")
%>
                <tr> 
                  <td width="20"><input name="grade" type="hidden" value="<%=grade%>"> 
                  </td>
                  <td><div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_unidade)%>
                      <input name="unidade" type="hidden" id="unidade" value="<% = unidade %>">
                      </font></div></td>
                  <td><div align="center"> <font class="form_dado_texto"> 
                      <%
response.Write(no_curso)%>
                      <input type="hidden" name="curso" value="<% = curso %>">
                      </font></div></td>
                  <td><div align="center"> <font class="form_dado_texto"> 
                      <%

		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' and CO_Curso ='"& curso &"'"  
		RS3.Open SQL3, CON0
		
if RS3.EOF THEN
no_etapa="sem etapa"
else
no_etapa=RS3("NO_Etapa")
end if
response.Write(no_etapa)%>
                      <input name="etapa" type="hidden" id="etapa" value="<% = co_etapa %>">
                      </font></div></td>
                  <td width="39"><div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(turma)%>
                      </font><font class="form_dado_texto"> 
                      <input name="turma" type="hidden" id="turma" value="<% = turma%>">
                      </font></div></td>
                  <td><div align="center"><font class="form_dado_texto"> 
                      <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Materia where CO_Materia ='"& mat_prin &"'" 
		RS4.Open SQL4, CON0
		
if RS4.EOF THEN
no_mat_prin="sem disciplina"
else
no_mat_prin=RS4("NO_Materia")
end if
response.Write(no_mat_prin)%>
                      </font><font class="form_dado_texto"> 
                      <input name="mat_prin" type="hidden" id="mat_prin" value="<% = mat_prin%>">
                      </font> </div></td>
                  <td><div align="center"><font class="form_dado_texto"> 
                      <%
select case tabela
case "TB_NOTA_A" 
response.Write("Modelo A")
case "TB_NOTA_B" 
response.Write("Modelo B")
case "TB_NOTA_C"
response.Write("Modelo C")
end select

%>
                      </font><font class="form_dado_texto"> 
                      <input name="tabela" type="hidden" id="tabela" value="<% = tabela%>">
                      </font> </div></td>
                  <td><div align="center"><font class="form_dado_texto"> 
                      <%
					  
					  
		Set CONu = Server.CreateObject("ADODB.Connection") 
		ABRIRu = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONu.Open ABRIRu
							  
		Set RS8 = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT * FROM TB_Usuario where CO_Usuario ="& coordenador
		RS8.Open SQL8, CONu
		
no_coordenador = RS8("NO_Usuario")
							  
					  response.Write(no_coordenador)%>
                      </font><font class="form_dado_texto"> 
                      <input name="coordenador" type="hidden" id="coordenador" value="<% = coordenador%>">
                      </font> </div></td>
                </tr>
              </table></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td bgcolor="#FFFFFF"> <table width="500" border="0" align="center" cellspacing="0">
                <tr> 
                  <td width="25%"> <div align="center"> 
                      <input name="alterar" type="button" class="botao_prosseguir" id="alterar" onClick="MM_goToURL('self','bd2.asp?opt=no&or=02');return document.MM_returnValue" value="N&atilde;o Apagar Notas">
                    </div></td>
                  <td width="25%"><div align="center"> 
                      <input name="Submit" type="submit" class="botao_prosseguir" value="Apagar Notas">
                    </div></td>
                </tr>
              </table></td>
          </tr>
        </table>
      </form></td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
</body>
</html>

<%



else
'response.Write(i&"<br")
pendentes2= pendentes2&", "&unidade&"-"&curso&"-"&co_etapa&"-"&turma&"-"&mat_prin&"-"&mat_fil&"-"&tabela&"-"&coordenador
end if

next
session("pendentes")=pendentes2


elseif opt="de" then

for i =1 to ubound(vertorExclui)
exclui = split(vertorExclui(i),"-")
unidade = exclui(0)
curso= exclui(1)
co_etapa= exclui(2)
turma= exclui(3)
mat_prin= exclui(4)
mat_fil= exclui(5)
tabela= exclui(6)
coordenador= exclui(7)


if i=1 then

if tabela ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na

elseif tabela="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb
elseif tabela ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc
end if	

		Set CON_A = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_A.Open ABRIR
		
		Set CON_N = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIR3
		
		Set RSa = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RSa = CON_A.Execute(SQL_A)

While Not RSa.EOF
nu_matricula = RSa("CO_Matricula")

	  	Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL_N = "Delete * from "& tabela &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia_Principal='"& mat_fil &"'AND CO_Materia='"& mat_prin &"'"
		Set RS3 = CON_N.Execute(SQL_N)

RSa.MOVENEXT
WEND

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "DELETE * from TB_Da_Aula where CO_Professor="& cod_prof &" AND CO_Materia_Principal='"& mat_prin &"'AND CO_Materia='"& mat_fil &"'AND NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		RS.Open SQL, CON_g
		
cod_prof = cod_prof

call GravaLog (chave,"PROF:"&cod_prof&"/U:"&unidade&"/C:"&curso&"/E:"&co_etapa&"/T:"&tabela&"/D:"&mat_fil)
else
pendentes2= pendentes2&", "&unidade&"-"&curso&"-"&co_etapa&"-"&turma&"-"&mat_prin&"-"&mat_fil&"-"&tabela&"-"&coordenador
end if
next
session("pendentes")=pendentes2
response.Redirect("bd2.asp?opt=sw")


elseif opt="no" then
for i=1 to ubound(vertorExclui)
exclui = split(vertorExclui(i),"-")
unidade = exclui(0)
curso= exclui(1)
co_etapa= exclui(2)
turma= exclui(3)
mat_prin= exclui(4)
mat_fil= exclui(5)
tabela= exclui(6)
coordenador= exclui(7)
if i=1 then
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "DELETE * from TB_Da_Aula where CO_Professor="& cod_prof &" AND CO_Materia_Principal='"& mat_prin &"'AND CO_Materia='"& mat_fil &"'AND NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		RS.Open SQL, CON_g
		
cod_prof = cod_prof

call GravaLog (chave,"PROF:"&cod_prof&"/U:"&unidade&"/C:"&curso&"/E:"&co_etapa&"/T:"&tabela&"/D:"&mat_fil)
else
pendentes2= pendentes2&", "&unidade&"-"&curso&"-"&co_etapa&"-"&turma&"-"&mat_prin&"-"&mat_fil&"-"&tabela&"-"&coordenador
end if
next
session("pendentes")=pendentes2
response.Redirect("bd2.asp?opt=sw&or=02")

end if
end if

%>
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
