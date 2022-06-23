<%'On Error Resume Next%>
<% Response.Charset="ISO-8859-1" %>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%
opt= request.QueryString("opt")
ori= request.QueryString("ori")

cod_cons= request.QueryString("cod_cons")


nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo")
ano_letivo_real = ano_letivo
sistema_local=session("sistema_local")

nvg=session("nvg")
session("nvg")=nvg
ano_info=nivel&"-"&nvg&"-"&ano_letivo
	
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano_info=nivel&"-"&nvg&"-"&ano_letivo

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


call navegacao (CON,nvg,nivel)
navega=Session("caminho")	


Set RS = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod_cons
RS.Open SQL, CON1


cod_cons = RS("CO_Matricula")
nome_aluno= RS("NO_Aluno")

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


 %> 
 
<tr>

            <td valign="top"> 
<FORM name="formulario" METHOD="POST" ACTION="bd.asp">
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
                  <td width="150"  height="10"><div align="right"><font class="form_dado_texto"> Matr&iacute;cula: </font></div></td>
                  <td width="50"  height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; </font><font class="form_corpo">
                    <input name="cod_cons" type="hidden" id="cod_cons" value="<%=cod_cons%>">
                    <%response.Write(cod_cons)%>
                  </font></td>
                  <td width="150" height="10"><div align="right"><font class="form_dado_texto"> Nome: </font></div></td>
                  <td width="500"  height="10" ><font class="form_corpo">
                    <%response.Write(nome_aluno)%>
                  </font></td>
                  <td width="150" height="10" >&nbsp;</td>
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
                      <tr class="tb_subtit"> 
                        <td width="33%"><div align="right"><font class="form_dado_texto">Documento 
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
                        <td width="33%"><div align="right"><font class="form_corpo"> 
                            <%response.Write(no_doc_mat)%>
                            </font></div></td>
                        <td width="34%"><div align="center"> 
                            <table width="50%" border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr> 
                                <td width="8%"><input name="<%response.Write(co_doc_mat)%>" type="radio" value="S"></td>
                                <td width="38%"><font class="form_dado_texto"> Entregue 
                                  </font></td>
                                <td width="7%"><input type="radio" name="<%response.Write(co_doc_mat)%>" value="N" checked></td>
                                <td width="47%"><font class="form_dado_texto"> Pendente 
                                  </font></td>
                              </tr>
                            </table>
                          </div></td>
                        <td width="33%"><div align="center"><font class="form_dado_texto"> 
                            </font></div></td>
                      </tr>
                      <%else
data_ent=RSde("DA_Entrega_Documento")
%>
                      <tr> 
                        <td width="33%"><div align="right"><font class="form_corpo"> 
                            <%response.Write(no_doc_mat)%>
                            </font></div></td>
                        <td width="34%"><div align="center"> 
                            <table width="50%" border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr> 
                                <td width="8%"><input name="<%response.Write(co_doc_mat)%>" type="radio" value="S" checked></td>
                                <td width="38%"><font class="form_dado_texto"> Entregue 
                                  </font></td>
                                <td width="7%"><input type="radio" name="<%response.Write(co_doc_mat)%>" value="N"></td>
                                <td width="47%"><font class="form_dado_texto"> Pendente 
                                  </font></td>
                              </tr>
                            </table>
                          </div></td>
                        <td width="33%"><div align="center" class="form_dado_texto">
                            <%response.Write(data_ent)%>
                            </div></td>
                      </tr>
                      <%
end if
RSdt.Movenext
wend
%>
                      <tr>
                        <td colspan="3"><hr></td>
                      </tr>
                      <tr>
                        <td><div align="center">
                            <input name="SUBMIT5" type=button class="botao_cancelar" onClick="MM_goToURL('parent','index.asp?nvg=WS-MA-MA-CDE');return document.MM_returnValue" value="Voltar">
						  </div></td>
                        <td>&nbsp;</td>
                        <td><div align="center"> 
                            <input name="SUBMIT" type=SUBMIT class="botao_prosseguir" value="Confirmar">
                        </div></td>
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
		  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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