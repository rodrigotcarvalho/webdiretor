<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<% 

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
cod= request.QueryString("cod_cons")
opt = request.QueryString("opt")
	

obr=cod


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")		

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
	
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
response.redirect("index.asp?nvg=WS-MA-MA-INA&opt=e1")

else
ano_aluno = RS("NU_Ano")
situacao = RS("CO_Situacao")
encerramento= RS("DA_Encerramento")
unidade= RS("NU_Unidade")
curso= RS("CO_Curso")
etapa= RS("CO_Etapa")
turma= RS("CO_Turma")
cham= RS("NU_Chamada")
motivo= RS("DS_Motivo")



Call LimpaVetor2

call GeraNomes("PORT",unidade,curso,etapa,CON0)
no_unidade = session("no_unidades")
no_curso = session("no_grau")
no_etapa = session("no_serie")

			ano = DatePart("yyyy", now)
			mes = DatePart("m", now) 
			dia = DatePart("d", now) 
			hora = DatePart("h", now) 
			min = DatePart("n", now) 

 			if situacao="L" or situacao="E" or situacao="R" then
			data_exibe=encerramento			
			else
			data_exibe=dia&"/"&mes&"/"&ano
			end if
end if
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

function MM_showHideLayers() { //v6.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
    obj.visibility=v; }
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function checksubmit()
{
  if (document.busca.situacao.value == 0)
  {    alert("Por favor selecione uma situação de inativação para o aluno!")
    document.busca.situacao.focus()
    return false
  }
    if (document.busca.motivo.value == "")
  {    alert("Por favor digite uma motivo para a inativação deste aluno!")
    document.busca.motivo.focus()
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

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</head>
<% if opt="listall" or opt="list" then%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%else %>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%end if %>
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
 <%if opt="ok" then%>
             <tr> 
         
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,9705,2,0) %>
    </td>
			  </tr>
 <%end if%>		  	  
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,317,0,0) %>
    </td>
			  </tr>			  
        <form action="bd.asp" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
          <tr>
      <td valign="top"> 
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo" height="177"
>
          <tr> 
            <td width="1000" class="tb_tit" height="15"
>Dados Escolares</td>
            <td width="218" class="tb_tit" height="15"
> </td>
          </tr>
          <tr> 
            <td height="21" width="1000"> <table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="19%" height="10"> <div align="right"><font class="form_dado_texto"> 
                      Matr&iacute;cula: </font></div></td>
                  <td width="9%" height="10"><font class="form_dado_texto"> 
                    <input name="cod" type="hidden" value="<%=codigo%>">
                    <%response.Write(codigo)%>
                    </font></td>
                  <td width="6%" height="10"> <div align="right"><font class="form_dado_texto"> 
                      Nome: </font></div></td>
                  <td width="66%" height="10"><font class="form_dado_texto"> 
                    <%response.Write(nome_prof)%>
                    </font></td>
                </tr>
              </table></td>
            <td valign="top" width="218" height="21"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font> </td>
          </tr>
          <tr> 
            <td height="19" bgcolor="#FFFFFF" width="1000">&nbsp;</td>
            <td valign="top" bgcolor="#FFFFFF" width="218" height="19">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="2" width="1000" height="32"><table width="100%" border="0" cellspacing="0">
                <tr class="tb_subtit"> 
                  <td width="113" height="10"> <div align="center">Unidade</div></td>
                  <td width="133" height="10"> <div align="center">Curso</div></td>
                  <td width="85" height="10"> <div align="center"> Etapa</div></td>
                  <td width="90" height="10"> <div align="center">Turma </div></td>
                  <td width="54" height="10"> <div align="center">Chamada</div></td>
                  <td width="86"> <div align="center"> Situa&ccedil;&atilde;o</div></td>
                </tr>
                <tr class="tb_corpo"> 
                  <td width="113" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_unidade)%>
                      </font></div></td>
                  <td width="133" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_curso)%>
                      </font></div></td>
                  <td width="85" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_etapa)%>
                      </font></div></td>
                  <td width="90" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(turma)%>
                      </font></div></td>
                  <td width="54" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(cham)%>
                      </font></div></td>
                  <td width="86"> <div align="center"> <font class="form_dado_texto"> 
                      <select name="situacao" class="borda" id="situacao">
                        <option value="0"></option>
                        <%
					
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno order by TX_Descricao_Situacao"
		RSCONTST.Open SQLCONTST, CON0

		while not RSCONTST.EOF
				tp_situacao =RSCONTST("CO_Situacao")
				no_situacao = RSCONTST("TX_Descricao_Situacao")	
 			if tp_situacao="L" or tp_situacao="E" or tp_situacao="R" then
					 if tp_situacao=situacao then%>
                        <option value="<%response.Write(tp_situacao)%>" selected> 
                        <%response.Write(no_situacao)%>
                        </option>
                        <%else%>
                        <option value="<%response.Write(tp_situacao)%>"> 
                        <%response.Write(no_situacao)%>
                        </option>
                        <%end if
			RSCONTST.MOVENEXT
			else
			RSCONTST.MOVENEXT
			end if
		WEND	
					%>
                      </select>
                      </font></div></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" width="1000" height="19">&nbsp;</td>
            <td width="218" height="19">&nbsp;</td>
          </tr>
          <tr valign="top"> 
            <td width="1000" height="48" colspan="2"> <table width="1000" border="0" cellspacing="0" cellpadding="0" height="46">
                <tr class="tb_subtit"> 
                  <td width="201" height="13"><div align="center">Data de Inativa&ccedil;&atilde;o</div></td>
                  <td width="799" height="13">Motivo</td>
                </tr>
                <tr class="<%=cor%>"> 
                  <td width="201" height="23" valign="top"> <div align="center"> 
                      <font class="form_dado_texto"> 
                      <%response.Write(data_exibe)%>
                      </font> </div></td>
                  <td width="799" height="23"> <font class="form_dado_texto"> 
                    <textarea name="motivo" cols="150" rows="5" wrap="VIRTUAL" id="motivo"><%response.Write(motivo)%></textarea>
                    </font></td>
                </tr>
              </table></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="1000" height="23" colspan="2"> <div align="center"> 
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr class="tb_corpo"> 
                    <td colspan="3"><hr></td>
                  </tr>
                  <tr> 
                    <td width="33%"><div align="center"> 
                        <input type="button" name="Submit2" value="Voltar" class="borda_bot3" onClick="MM_goToURL('parent','index.asp?nvg=WS-MA-MA-INA')">
                      </div></td>
                    <td width="34%">&nbsp;</td>
                    <td width="33%"> <div align="center"> 
                       <input type="submit" name="Submit" value="Confirmar" class="borda_bot">
                      </div></td>
                  </tr>
                </table>
                <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> </font></div></td>
          </tr>
        </table></td>
    </tr>
</form>
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

</body>
<script type="text/javascript">
<!--
  initInputHighlightScript();
//-->
</script>


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