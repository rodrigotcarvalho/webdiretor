<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
          <% 

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

ori = request.QueryString("ori")
opt= request.QueryString("opt")

chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
	

if ori="prd" then
	periodo_check=request.form("periodo")
	cod= request.QueryString("cod_cons")
elseif opt="ok" then
	cod= request.QueryString("cod_cons")
	periodo_check = request.QueryString("per")
else
	cod= request.QueryString("cod_cons")
	periodo_check=1
end if
obr=cod&"?"&periodo_check


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

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod
		RS.Open SQL, CON1
		
		
codigo = RS("CO_Matricula")
nome_prof = RS("NO_Aluno")



		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod
		RS.Open SQL, CON1

if RS.EOF then
	response.Write("ERRO - Aluno "&cod&" não matriculado no ano letivo "& ano_letivo)
	response.End()
else
	ano_aluno = RS("NU_Ano")
	rematricula = RS("DA_Rematricula")
	situacao = RS("CO_Situacao")
	encerramento= RS("DA_Encerramento")
	unidade= RS("NU_Unidade")
	curso= RS("CO_Curso")
	etapa= RS("CO_Etapa")
	turma= RS("CO_Turma")
	cham= RS("NU_Chamada")
end if

Call LimpaVetor2

%>
          <html>
          <head>
          <title>Web Diretor</title>
          <link href="../../../../estilos.css" rel="stylesheet" type="text/css">
          <script language="JavaScript" type="text/JavaScript">
<!--

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
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
                </font></td>
            </tr>
            <%if opt = "ok" then%>
            <tr>
              <td height="10" valign="top"><%
		call mensagens(nivel,622,2,0)
%>
                <div align="center"></div></td>
            </tr>
            <%elseif opt= "err6" then %>
            <tr>
              <td height="10" valign="top"><%
	call mensagens(nivel,620,1,errante)
%></td>
            </tr>
            <%end if
%>
            <tr>
              <td height="10" colspan="5" valign="top"><%call mensagens(nivel,636,0,0) %></td>
            </tr>
            <tr>
              <td valign="top"><table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
                  <tr>
                  <td width="653" class="tb_tit"
>Dados Escolares</td>
                  <td width="113" class="tb_tit"
></td>
                </tr>
                  <tr>
                  <td height="10"><table width="100%" border="0" cellspacing="0">
                      <tr>
                      <td width="19%" height="10"><div align="right"><font class="form_dado_texto"> Matr&iacute;cula: </font></div></td>
                      <td width="9%" height="10"><font class="form_dado_texto">
                        <%response.Write(codigo)%>
                        </font></td>
                      <td width="6%" height="10"><div align="right"><font class="form_dado_texto"> Nome: </font></div></td>
                      <td width="66%" height="10"><font class="form_dado_texto">
                        <%response.Write(nome_prof)%>
                        &nbsp;</font></td>
                    </tr>
                    </table></td>
                  <td valign="top">&nbsp;</td>
                </tr>
                  <tr>
                  <td height="10" bgcolor="#FFFFFF">&nbsp;</td>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                  <tr>
                  <td colspan="2"><form action="altera.asp?ori=prd&cod_cons=<%response.Write(codigo)%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
                      <table width="100%" border="0" cellspacing="0">
                      <tr class="tb_subtit">
                          <td width="40" height="10"><div align="center">
                              <%
call GeraNomes("PORT",unidade,curso,etapa,CON0)
no_unidades = session("no_unidades")
no_grau = session("no_grau")
no_serie = session("no_serie")
%>
                              Ano</div></td>
                          <td width="80" height="10"><div align="center">Matr&iacute;cula</div></td>
                          <td width="100" height="10"><div align="center">Cancelamento</div></td>
                          <td width="100" height="10"><div align="center"> Situa&ccedil;&atilde;o</div></td>
                          <td width="100" height="10"><div align="center">Unidade</div></td>
                          <td width="130" height="10"><div align="center">Curso</div></td>
                          <td width="100" height="10"><div align="center"> Etapa</div></td>
                          <td width="100" height="10"><div align="center">Turma </div></td>
                          <td width="100" height="10"><div align="center">Chamada</div></td>
                          <td width="150"><div align="center">Per&iacute;odo</div></td>
                        </tr>
                      <tr class="tb_corpo"
>
                          <td width="40" height="10"><div align="center"> <font class="form_dado_texto">
                            <%response.Write(ano_aluno)%>
                            </font></div></td>
                          <td width="80" height="10"><div align="center"> <font class="form_dado_texto">
                            <%response.Write(rematricula)%>
                            </font></div></td>
                          <td width="100" height="10"><div align="center"> <font class="form_dado_texto">
                            <%response.Write(encerramento)%>
                            </font></div></td>
                          <td width="100" height="10"><div align="center"> <font class="form_dado_texto">
                            <%
					
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
		RSCONTST.Open SQLCONTST, CON0
							
				no_situacao = RSCONTST("TX_Descricao_Situacao")	
					response.Write(no_situacao)%>
                            </font></div></td>
                          <td width="100" height="10"><div align="center"> <font class="form_dado_texto">
                            <%response.Write(no_unidades)%>
                            </font></div></td>
                          <td width="130" height="10"><div align="center"> <font class="form_dado_texto">
                            <%response.Write(no_grau)%>
                            </font></div></td>
                          <td width="100" height="10"><div align="center"> <font class="form_dado_texto">
                            <%response.Write(no_serie)%>
                            </font></div></td>
                          <td width="100" height="10"><div align="center"> <font class="form_dado_texto">
                            <%response.Write(turma)%>
                            </font></div></td>
                          <td width="100" height="10"><div align="center"> <font class="form_dado_texto">
                            <%response.Write(cham)%>
                            </font></div></td>
                          <td width="150"><div align="center">
                              <select name="periodo" class="select_style" id="periodo" onChange="MM_callJS('submitfuncao()')">
                              <%
		Set RSPER = Server.CreateObject("ADODB.Recordset")
		SQLPER = "SELECT * FROM TB_Periodo order by NU_Periodo"'"
		RSPER.Open SQLPER, CON0
		
		While not RSPER.EOF
		periodo=RSPER("NU_Periodo")
		no_periodo=RSPER("NO_Periodo")
		periodo=periodo*1
		periodo_check=periodo_check*1
		
			if periodo=periodo_check then		
			%>
                              <option value="<%=periodo%>" selected><%=no_periodo%></option>
                              <%else%>
                              <option value="<%=periodo%>"><%=no_periodo%></option>
                              <%end if
		RSPER.Movenext
		WEND
		%>
                            </select>
                            </div></td>
                        </tr>
                    </table>
                    </form></td>
                </tr>
                  <tr>
                  <td bgcolor="#FFFFFF">&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                  <tr>
                  <td height="10" colspan="2" class="tb_tit"
>Avalia&ccedil;&otilde;es</td>
                </tr>
                  <tr>
                  <td colspan="2"></td>
                </tr>
                  <tr bgcolor="#FFFFFF">
                  <td colspan="2" valign="top"><table width="100%" border="0" align="center" cellspacing="0" class="tb_corpo">
                      <tr>
                      <td valign="top"><%		Set RS_tb = Server.CreateObject("ADODB.Recordset")
		SQL_tb = "SELECT * FROM TB_Da_Aula WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso&"' AND CO_Etapa ='"& etapa&"' AND CO_Turma ='"& turma&"'"
		RS_tb.Open SQL_tb, CON2

if RS_tb.eof then
%>
                          <div align="center"><font class="form_dado_texto">
                            <%response.Write("<br><br><br><br><br>N&atilde;o existem Avalia&ccedil;&otilde;es para este aluno!")%>
                            </font></div>
                          <%
	lancar = "N"						  
else
notaFIL=RS_tb("TP_Nota")
periodo=periodo_check 
	lancar = "S"						  
	if notaFIL ="TB_NOTA_A" then
		CAMINHOn = CAMINHO_na
		opcao="A"
	elseif notaFIL="TB_NOTA_B" then
		CAMINHOn = CAMINHO_nb
		opcao="B"	
	elseif notaFIL ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc
		opcao="C"			
	elseif notaFIL ="TB_NOTA_D" then
		CAMINHOn = CAMINHO_nd
		opcao="D"	
	elseif notaFIL ="TB_NOTA_E" then
		CAMINHOn = CAMINHO_ne	
		opcao="E"					
	elseif notaFIL ="TB_NOTA_F" then
		CAMINHOn = CAMINHO_nf	
		opcao="F"					
	elseif notaFIL ="TB_NOTA_V" then
		CAMINHOn = CAMINHO_nv	
		opcao="V"					
	elseif notaFIL ="TB_NOTA_K" then
		CAMINHOn = CAMINHO_nk	
		opcao="K"					
	else
		response.Write("ERRO")
	end if
end if

if	lancar = "S" then

%>
                          
                          <!--#include file="../../../../inc/lanca_notas_aluno.asp"-->
                          
<%end if%>                          
                          </td>
                    </tr>
                    </table></td>
                </tr>
                  <tr class="tb_tit"
>
                  <td colspan="2">&nbsp;</td>
                </tr>
                </table></td>
            </tr>
            <tr>
              <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
            </tr>
          </table>
          </body>
          <%call GravaLog (chave,obr)%>
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