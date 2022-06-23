<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->
<%
nivel=4

opt=request.QueryString("opt")
cod_cons= request.QueryString("cod_cons")
autoriza=Session("autoriza")
Session("autoriza")=autoriza

permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("nvg")
session("nvg")=chave
session("chave")=session("nvg")
ano_info=nivel&"-"&chave&"-"&ano_letivo

vetor_sigla_periodo=periodos(periodo, "todas_siglas")
vetor_cod_periodo=periodos(periodo, "todos_codigos")

siglas_periodos=split(vetor_sigla_periodo,"#!#")
vetor_periodo=split(vetor_cod_periodo,"#!#")

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
		
		Set RSp = Server.CreateObject("ADODB.Recordset")
		SQLp = "SELECT * FROM TB_Professor WHERE CO_Professor ="& cod_cons
		RSp.Open SQLp, CONp

cod_cons = RSp("CO_Professor")
nome_prof = RSp("NO_Professor")
co_usr_prof = RSp("CO_Usuario")

 call navegacao (CON,chave,nivel)
navega=Session("caminho")

call linkFuncao(CON,"WA","PF","CN","MNL",nivel)
link_funcao=session("link_funcao")

	%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../js/global.js"></script>
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
function submitfuncao()  
{
   var f=document.forms[3]; 
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
<%call cabecalho(nivel)
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
   <tr>                   
    <td height="10"> 
      <%	call mensagens(4,609,0,0) 

%>
    </td>
                </tr>				  				  


          <tr> 
            <td valign="top"> 
             
                
        <table width="1000" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
        <tr> 
                    
            <td> <form name="alteracao" method="post" action="grade_cp2.asp">
                
              <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
                <tr> 
                    
                  <td width="653" height="15" class="tb_tit"> Professor</td>
                  </tr>
                  <tr> 
                    <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                          <td width="9%" height="30" class="tb_subtit"> <div align="right">C&oacute;digo: 
                              </div></td>
                          <td width="11%" height="30"><font class="form_dado_texto"> 
                            <input name="cod_cons" type="hidden" value="<%=cod_cons%>">
                            <%response.Write(cod_cons)%>
                            <input name="tp" type="hidden" id="tp" value="P">
                            <input name="acesso" type="hidden" id="acesso" value="2">
                            <input name="nome_prof" type="hidden" id="nome_prof" value="<% =nome_prof%>">
                            <input name="co_usr_prof" type="hidden" id="co_usr_prof" value="<% =co_usr_prof%>">
                            </font></td>
                          <td width="6%" height="30" class="tb_subtit"> <div align="right">Nome: </div>
                            </td>
                          <td width="74%" height="30"><font class="form_dado_texto"> 
                            <%response.Write(nome_prof)%>
                            </font> </td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr> 
                    
                  <td height="15" class="tb_tit"> Grade de Aulas</td>
                  </tr>
                  <tr> 
                    <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="130" class="tb_subtit"> <div align="center">UNIDADE 
                        </div></td>
                        <td width="130" class="tb_subtit"> <div align="center">CURSO 
                          </div></td>
                        <td width="130" class="tb_subtit"> <div align="center">ETAPA 
                          </div></td>
                        <td width="130" class="tb_subtit"> <div align="center">TURMA 
                          </div></td>
                        <td width="240" class="tb_subtit"> 
                          <div align="center">DISCIPLINA</div></td>
                      <%				  
						for sp=0 to ubound(siglas_periodos)
						%>
                        <td width="50" class="tb_subtit"> <div align="center"><strong><%response.Write(siglas_periodos(sp))%></strong></div></td>
						<%next%>
                      </tr>
                      <%
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Da_Aula where CO_Professor ="& cod_cons 
		RS1.Open SQL1, CONG
		

if RS1.EOF THEN
ELSE
check=2
while not RS1.EOF
cod_cons = RS1("CO_Professor")
curso = RS1("CO_Curso")
unidade = RS1("NU_Unidade")
co_etapa= RS1("CO_Etapa")
turma= RS1("CO_Turma")
mat_prin = RS1("CO_Materia_Principal")
mat_fil = RS1("CO_Materia")
tabela = RS1("TP_Nota")
coordenador= RS1("CO_Cord")
		
		valor = unidade&"-"&curso&"-"&co_etapa&"-"&turma&"-"&mat_prin&"-"&mat_fil&"-"&tabela&"-"&coordenador


	Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RSu.Open SQLu, CON0
		
no_unidade = RSu("NO_Unidade")

		Set RSc = Server.CreateObject("ADODB.Recordset")
		SQLc = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RSc.Open SQLc, CON0
		
no_curso = RSc("NO_Abreviado_Curso")

 if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if
%>
                      <tr> 
                        <td width="130" class="<%=cor%>"> <div align="center"> 
                            <font class="form_dado_texto"> 
                            <%response.Write(no_unidade)%>
                            </font></div></td>
                        <td width="130" class="<%=cor%>"> <div align="center"> 
                            <font class="form_dado_texto"> 
                            <%
response.Write(no_curso)%>
                            </font></div></td>
                        <td width="130" class="<%=cor%>"> <div align="center"> 
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
                        <td width="130" class="<%=cor%>"> <div align="center"> 
                            <font class="form_dado_texto"> 
                            <%response.Write(turma)%>
                            </font></div></td>
                        <td width="240" class="<%=cor%>"> 
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
                            </font> </div></td>
<%     for qp=0 to ubound(vetor_periodo) %>                            
                        <td width="50" class="<%=cor%>"> <div align="center"> 
                            <%					

		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Da_Aula where CO_Professor="& cod_cons &"AND CO_Materia_Principal='"& mat_prin &"'AND NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		RS4.Open SQL4, CONG



                if RS4.EOF then
                
                else
                    periodo=vetor_periodo(qp)
                    p = RS4("ST_Per_"&periodo)
        
                    if p = "x" then
                        img="s.gif"
                    else
                        img="n.gif"
                    end if	
                    %>
                      <div align="center"><a href="notas.asp?d=<%response.Write(co_mat)%>&pr=<%response.Write(co_professor)%>&p=<%response.Write(periodo)%>"><img src="../../../../img/<%response.Write(img)%>" width="8" height="8" border="0"></a></div>
                      <%
                end if%>
                          </div></td>
<%next%>  
                        
                      </tr>
                      <%
check=check+1
RS1.MOVENEXT
WEND
END IF				
%>
                      <tr> 
                        <td width="130" height="15"></td>
                        <td width="130" height="15"></td>
                        <td width="130" height="15"></td>
                        <td width="130"> </td>
                        <td width="240"> </td>
                        <td width="50"> </td>
                      </tr>
                      <tr class="tb_tit"> 
                        <td width="130"></td>
                        <td width="130"></td>
                        <td width="130"></td>
                        <td width="130">&nbsp;</td>
                        <td width="240">&nbsp;</td>
                      <%				  
						for sp=0 to ubound(siglas_periodos)
						%>
                        <td width="50">&nbsp;</strong></div></td>
						<%next%>
                      </tr>
                    </table></td>
                  </tr>
                </TABLE>
              </form></td>
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