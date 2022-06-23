<%On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->


<!--#include file="../../../../inc/caminhos.asp"-->



<% 
Session.LCID = 1046
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
ori = request.QueryString("opt")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

trava=session("trava")

dados_ocor= split(ori,"?")
cod=dados_ocor(0)
da_ocorrencia=dados_ocor(1)
ho_ocorrencia=dados_ocor(2)
co_ocorrencia=dados_ocor(3)




tp_ocor=Session("tp_ocor")
data_de=Session("data_de")
hora_de=Session("hora_de")
data_inicio=Session("data_inicio")
data_ate=Session("data_ate")
hora_ate=Session("hora_ate")
data_fim=Session("data_fim")

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

		Set CON3 = Server.CreateObject("ADODB.Connection") 
		ABRIR3 = "DBQ="& CAMINHO_o & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3
		
		
		Set CONp = Server.CreateObject("ADODB.Connection") 
		ABRIRp = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONp.Open ABRIRp		
		
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


ano_aluno = RS("NU_Ano")
rematricula = RS("DA_Rematricula")
situacao = RS("CO_Situacao")
encerramento= RS("DA_Encerramento")
unidade= RS("NU_Unidade")
curso= RS("CO_Curso")
etapa= RS("CO_Etapa")
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


<!--
var checkflag = "false";
function check(field) {
if (checkflag == "false") {
for (i = 0; i < field.length; i++) {
field[i].checked = true;}
checkflag = "true";
return "Desmarcar Todos"; }
else {
for (i = 0; i < field.length; i++) {
field[i].checked = false; }
checkflag = "false";
return "Marcar Todos"; }
}
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
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--


function checksubmit()
{
  if (document.busca.tp_ocor.value == "999999")
  {    alert("Por favor selecione um tipo de ocorrência!")
    document.busca.tp_ocor.focus()
    return false
 }aula = document.busca.aula.value;
     if (aula.length > 3)
  {    alert("O valor do campo Aula deve possuir menos que 3 caracteres")
    document.busca.aula.focus()
    return false
  }
//    if (document.busca.observacao.value == "")
//  {    alert("Por favor digite uma observação!")
//    document.busca.observacao.focus()
//    return false
//  }
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
      <%call mensagens(nivel,316,0,0) %>
    </td>
			  </tr>
<% IF trava="n" then %>
        <form action="bd.asp?opt=alt" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
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
                    <input name="nome" type="hidden" class="textInput" id="nome2"  value="<%response.Write(nome_prof)%>" size="75" maxlength="50">
                    <input name="assunto" type="hidden" class="textInput" id="nome"  value="PED" size="75" maxlength="50">
                    </font></td>
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
	no_unidade= GeraNomes("U",unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)
	no_curso=GeraNomes("C",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro) 	
	no_etapa=GeraNomes("E",curso,co_etapa,variavel3,variavel4,variavel5,CON0,outro) 	
%>
                      Ano</div></td>
                  <td width="81" height="10"> <div align="center">Matr&iacute;cula</div></td>
                  <td width="75" height="10"> <div align="center">Cancelamento</div></td>
                  <td width="86" height="10"> <div align="center"> Situa&ccedil;&atilde;o</div></td>
                  <td width="113" height="10"> <div align="center">Unidade</div></td>
                  <td width="133" height="10"> <div align="center">Curso</div></td>
                  <td width="85" height="10"> <div align="center"> Etapa</div></td>
                  <td width="90" height="10"> <div align="center">Turma </div></td>
                  <td width="54" height="10"> <div align="center">Chamada</div></td>
                </tr>
                <tr class="tb_corpo"> 
                  <td width="33" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(ano_aluno)%>
                      </font></div></td>
                  <td width="81" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(rematricula)%>
                      </font></div></td>
                  <td width="75" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(encerramento)%>
                      </font></div></td>
                  <td width="86" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%
					
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
		RSCONTST.Open SQLCONTST, CON0
							
				no_situacao = RSCONTST("TX_Descricao_Situacao")	
					response.Write(no_situacao)%>
                      </font></div></td>
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
                </tr>
                <tr class="tb_tit"> 
                  <td height="10" colspan="5">Ocorr&ecirc;ncia</td>
                  <td height="10" colspan="2">&nbsp;</td>
                  <td height="10" colspan="2">&nbsp;</td>
                </tr>
                <tr class="tb_subtit"> 
                  <td height="10" colspan="3">Ocorr&ecirc;ncia</td>
                  <td height="10" colspan="2">Professor:</td>
                  <td height="10"><div align="left">Disciplina</div></td>
                  <td height="10">Data </td>
                  <td height="10"> <div align="left">Hora</div></td>
                  <td height="10">Aula</td>
                </tr>
                <tr class="tb_corpo"> 
                  <td height="10" colspan="3"><div align="left"> <font class="form_dado_texto"> 
                      <%
				  
dados_data=split(da_ocorrencia,"/")
dia=dados_data(0)
mes=dados_data(1)
ano=dados_data(2)

dados_hora=split(ho_ocorrencia,":")
hora=dados_hora(0)
min=dados_hora(1)


'mid(ho_ocorrencia,1,10) 
da_ocorrencia_cons=mes&"/"&dia&"/"&ano

		'response.Write "SELECT * FROM TB_Ocorrencia_Aluno WHERE CO_Matricula ="& cod&" AND CO_Ocorrencia ="& co_ocorrencia &" AND DA_Ocorrencia= #"&da_ocorrencia_cons&"# AND mid(HO_Ocorrencia,1,16) =#12/30/1899 "&ho_ocorrencia&"#"

		Set RSo = Server.CreateObject("ADODB.Recordset")
		SQLo = "SELECT * FROM TB_Ocorrencia_Aluno WHERE CO_Matricula ="& cod&" AND CO_Ocorrencia ="& co_ocorrencia &" AND DA_Ocorrencia= #"&da_ocorrencia_cons&"# AND mid(HO_Ocorrencia,1,16)=#12/30/1899 "&ho_ocorrencia&"#"
		RSo.Open SQLo, CON3

co_assunto=RSo("CO_Assunto")
aula=RSo("NU_Aula")
co_professor=RSo("CO_Professor")
co_materia_original=RSo("NO_Materia")
observa=RSo("TX_Observa")
					  
 		Set RSto = Server.CreateObject("ADODB.Recordset")
		SQLto = "SELECT * FROM TB_Tipo_Ocorrencia where CO_Ocorrencia="&co_ocorrencia&" order by NO_Ocorrencia"
		RSto.Open SQLto, CON0
		

no_ocorrencia=RSto("NO_Ocorrencia")
					  				  
Response.Write(no_ocorrencia)%>
                      <input name="tp_ocor" type="hidden" id="tp_ocor" value="<%response.write(co_ocorrencia)%>">
                      </font></div></td>
                  <td height="10" colspan="2"> 
                    <%
 		Set RSmat = Server.CreateObject("ADODB.Recordset")
		SQLmat = "SELECT * FROM TB_Da_Aula Where NU_Unidade="&unidade&" AND CO_Curso='"&curso&"' AND CO_Etapa='"&etapa&"' AND CO_Turma='"&turma&"' order by CO_Materia_Principal"
		RSmat.Open SQLmat, CON2
prof_check="nada"
prof_qtd=0
co_materia_check="nada"
While not RSmat.EOF
co_materia=RSmat("CO_Materia_Principal")

 		Set RSnomat = Server.CreateObject("ADODB.Recordset")
		SQLnomat = "SELECT * FROM TB_Materia Where CO_Materia='"&co_materia&"'"
		RSnomat.Open SQLnomat, CON0

no_materia=RSnomat("NO_Materia")

prof=RSmat("CO_Professor")
if prof_check=prof then
count_prof=count_prof
else
prof_qtd=prof_qtd&"?"&prof
prof_check=prof
count_prof=count_prof+1
end if
if co_materia_check=co_materia then
RSmat.Movenext
else

co_materia_check=co_materia 
RSmat.Movenext
end if
WEND
 



If count_prof=1 then
If co_professor="" or isnull(co_professor) then					
else
 		Set RSpro = Server.CreateObject("ADODB.Recordset")
		SQLpro = "SELECT * FROM TB_Professor Where CO_Professor="&co_professor
		RSpro.Open SQLpro, CONp
prof=RSpro("CO_Professor")
no_prof=RSpro("NO_Professor")
response.Write(no_prof)
end if		
%></font>
                    <input name="no_prof" type="hidden" id="no_prof" value="<%response.Write(no_prof)%>"> 
                    <%else
If co_professor="" or isnull(co_professor) or co_professor="999999" or co_professor=999999 then					
else
 		Set RSpro = Server.CreateObject("ADODB.Recordset")
		SQLpro = "SELECT * FROM TB_Professor Where CO_Professor="&co_professor
		RSpro.Open SQLpro, CONp
no_prof_original=RSpro("NO_Professor")
end if		
dados= split(prof_qtd, "?" )
%>
                    <select name="no_prof" class="select_style" id="select2">
                      <option value="999999" selected></option>
                      <%
For i=1 to ubound(dados)

 
 		Set RSpro = Server.CreateObject("ADODB.Recordset")
		SQLpro = "SELECT * FROM TB_Professor Where CO_Professor="&dados(i)&" order by NO_Professor"
		RSpro.Open SQLpro, CONp


prof=RSpro("CO_Professor")
no_prof=RSpro("NO_Professor")
if no_prof=no_prof_original then
%>
                      <option value="<%=prof%>" selected> 
                      <%Response.Write(no_prof)%>
                      </option>
                      <%
else
%>
                      <option value="<%=prof%>"> 
                      <%Response.Write(no_prof)%>
                      </option>
                      <%
end if					  

next
%>
                    </select> 
                    <%end if%>
                  </td>
                  <td height="10"> <div align="center"><font class="form_dado_texto"> 
                      <select name="disciplina" class="select_style" id="select3">
                        <option value="999999" selected></option>
                        <%
 		Set RSmat = Server.CreateObject("ADODB.Recordset")
		SQLmat = "SELECT * FROM TB_Da_Aula Where NU_Unidade="&unidade&" AND CO_Curso='"&curso&"' AND CO_Etapa='"&etapa&"' AND CO_Turma='"&turma&"' order by CO_Materia_Principal"
		RSmat.Open SQLmat, CON2

co_materia_check="nada"
While not RSmat.EOF
co_materia=RSmat("CO_Materia_Principal")

 		Set RSnomat = Server.CreateObject("ADODB.Recordset")
		SQLnomat = "SELECT * FROM TB_Materia Where CO_Materia='"&co_materia&"'"
		RSnomat.Open SQLnomat, CON0

no_materia=RSnomat("NO_Materia")


if co_materia_check=co_materia then
RSmat.Movenext
else
if co_materia=co_materia_original then
%>
                        <option value="<%=co_materia%>" selected> 
                        <%Response.Write(no_materia)%>
                        </option>
                        <%
 else  
%>
                        <option value="<%=co_materia%>"> 
                        <%Response.Write(no_materia)%>
                        </option>
                        <%
end if												
co_materia_check=co_materia 
RSmat.Movenext
end if
WEND
%>
                      </select>
                      </font></div></td>
                  <td height="10"><font class="form_dado_texto"> 
                    <%
					  

dia=dia*1
mes=mes*1
hora=hora*1
min=min*1

if dia<10 then
dia="0"&dia
end if
if mes<10 then
mes="0"&mes
end if
if hora<10 then
hora="0"&hora
end if
if min<10 then
min="0"&min
end if
da_show=dia&"/"&mes&"/"&ano
data_altera=mes&"/"&dia&"/"&ano
hora_show=hora&":"&min

response.Write(da_show)					  
					  %>
                    <input name="data" type="hidden" id="data" value="<%=da_show%>">
                    <input name="data_altera" type="hidden" id="data_altera" value="<%=data_altera%>">
                    </font></td>
                  <td height="10"> <div align="left"><font class="form_dado_texto"> 
                      <%response.Write(hora_show)					  
					  %>
                      <input name="hora" type="hidden" id="hora" value="<%=hora_show%>">
                      </font> </div></td>
                  <td height="10"><input name="aula" type="text" class="textInput" id="aula3" value="<%response.write(aula)%>" size="15"></td>
                </tr>
                <tr class="tb_subtit"> 
                  <td height="10" colspan="3">Observa&ccedil;&atilde;o</td>
                  <td height="10">&nbsp;</td>
                  <td height="10" colspan="3">&nbsp;</td>
                  <td height="10">&nbsp;</td>
                  <td height="10">&nbsp;</td>
                </tr>
                <tr > 
                  <td height="10" colspan="9"><textarea name="observacao" cols="196" rows="5" wrap="VIRTUAL" id="observacao"><%response.Write(observa)%></textarea></td>
                </tr>
              </table></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></div></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2"><div align="center"> 
                <table width="1000" border="0" align="center" cellspacing="0">
                  <tr> 
                    <td colspan="3"><hr></td>
                  </tr>
                  <tr> 
                    <td width="33%"> 
                      <div align="center"> 
                        <input name="alterar" type="button" class="botao_cancelar" id="alterar" onClick="MM_goToURL('parent','resumo.asp?or=2');return document.MM_returnValue" value="Voltar">
                      </div></td>
                    <td width="34%">&nbsp;</td>
                    <td width="33%">
<div align="center"> 
                        <input name="Submit" type="submit" class="botao_prosseguir" value="Confirmar">
                      </div></td>
                  </tr>
                </table>
                <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> </font></div></td>
          </tr>
        </table></td>
    </tr>
</form>
<%else%>			  			  
        <form action="resumo.asp?or=3" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
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
                    <input name="nome" type="hidden" class="textInput" id="nome2"  value="<%response.Write(nome_prof)%>" size="75" maxlength="50">
                    <input name="tp_ocor" type="hidden" class="textInput" id="nome2"  value="<%response.Write(tp_ocor)%>" size="75" maxlength="50">
                    <input name="data_de" type="hidden" class="textInput" id="nome2"  value="<%response.Write(data_de)%>" size="75" maxlength="50">
                    <input name="hora_de" type="hidden" class="textInput" id="nome2"  value="<%response.Write(hora_de)%>" size="75" maxlength="50">
                    <input name="data_inicio" type="hidden" class="textInput" id="nome2"  value="<%response.Write(data_inicio)%>" size="75" maxlength="50">
                    <input name="data_ate" type="hidden" class="textInput" id="nome2"  value="<%response.Write(data_ate)%>" size="75" maxlength="50">
                    <input name="hora_ate" type="hidden" class="textInput" id="nome2"  value="<%response.Write(hora_ate)%>" size="75" maxlength="50">
                    <input name="data_fim" type="hidden" class="textInput" id="nome2"  value="<%response.Write(data_fim)%>" size="75" maxlength="50">
                    <input name="ordem" type="hidden" id="ordem" value="dt">
                    </font></td>
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
	no_unidade= GeraNomes("U",unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)
	no_curso=GeraNomes("C",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro) 	
	no_etapa=GeraNomes("E",etapa,variavel2,variavel3,variavel4,variavel5,CON0,outro) 		
%>
                      Ano</div></td>
                  <td width="81" height="10"> <div align="center">Matr&iacute;cula</div></td>
                  <td width="75" height="10"> <div align="center">Cancelamento</div></td>
                  <td width="86" height="10"> <div align="center"> Situa&ccedil;&atilde;o</div></td>
                  <td width="113" height="10"> <div align="center">Unidade</div></td>
                  <td width="133" height="10"> <div align="center">Curso</div></td>
                  <td width="85" height="10"> <div align="center"> Etapa</div></td>
                  <td width="90" height="10"> <div align="center">Turma </div></td>
                  <td width="54" height="10"> <div align="center">Chamada</div></td>
                </tr>
                <tr class="tb_corpo"> 
                  <td width="33" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(ano_aluno)%>
                      </font></div></td>
                  <td width="81" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(rematricula)%>
                      </font></div></td>
                  <td width="75" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(encerramento)%>
                      </font></div></td>
                  <td width="86" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%
					
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
		RSCONTST.Open SQLCONTST, CON0
							
				no_situacao = RSCONTST("TX_Descricao_Situacao")	
					response.Write(no_situacao)%>
                      </font></div></td>
                  <td width="113" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_unidades)%>
                      </font></div></td>
                  <td width="133" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_grau)%>
                      </font></div></td>
                  <td width="85" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_serie)%>
                      </font></div></td>
                  <td width="90" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(turma)%>
                      </font></div></td>
                  <td width="54" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(cham)%>
                      </font></div></td>
                </tr>
                <tr class="tb_tit"> 
                  <td height="10" colspan="5">Ocorr&ecirc;ncia</td>
                  <td height="10" colspan="2">&nbsp;</td>
                  <td height="10" colspan="2">&nbsp;</td>
                </tr>
                <tr class="tb_subtit"> 
                  <td height="10" colspan="5">Ocorr&ecirc;ncia</td>
                  <td height="10" colspan="2"><div align="center">Data</div></td>
                  <td height="10" colspan="2"> <div align="center">Hora</div></td>
                </tr>
                <tr class="tb_corpo"> 
                  <td height="10" colspan="5"><div align="left"> <font class="form_dado_texto"> 
                      <%
 		Set RSto = Server.CreateObject("ADODB.Recordset")
		SQLto = "SELECT * FROM TB_Tipo_Ocorrencia where CO_Ocorrencia="&co_ocorrencia&" order by NO_Ocorrencia"
		RSto.Open SQLto, CON0

no_ocorrencia=RSto("NO_Ocorrencia")
					  response.Write(no_ocorrencia)				  
					  %>
                      </font></div></td>
                  <td height="10" colspan="2"> <div align="center"><font class="form_dado_texto"> 
                      <%response.Write(da_ocorrencia)%>
                      </font></div></td>
                  <td height="10" colspan="2"> <div align="center"><font class="form_dado_texto"> 
                      <%response.Write(ho_ocorrencia)%>
                      </font></div></td>
                </tr>
                <tr class="tb_subtit"> 
                  <td height="10" colspan="2"><div align="left">Professor:</div></td>
                  <td height="10">&nbsp;</td>
                  <td height="10">Disciplina</td>
                  <td height="10" colspan="2">&nbsp;</td>
                  <td height="10">&nbsp;</td>
                  <td height="10">Aula</td>
                  <td height="10">&nbsp;</td>
                </tr>
                <%
dados_data=split(da_ocorrencia,"/")
dia_data=dados_data(0)
mes_data=dados_data(1)
ano_data=dados_data(2)

da_ocorrencia_cons=mes_data&"/"&dia_data&"/"&ano_data

		Set RSo = Server.CreateObject("ADODB.Recordset")
		SQLo = "SELECT * FROM TB_Ocorrencia_Aluno WHERE CO_Matricula ="& cod&" AND CO_Ocorrencia ="& co_ocorrencia &" AND DA_Ocorrencia= #"&da_ocorrencia_cons&"# AND HO_Ocorrencia=#12/30/1899 "&ho_ocorrencia&"#"
		RSo.Open SQLo, CON3

co_assunto=RSo("CO_Assunto")
aula=RSo("NU_Aula")
co_professor=RSo("CO_Professor")
co_materia=RSo("NO_Materia")
observa=RSo("TX_Observa")

 		Set RSnomat = Server.CreateObject("ADODB.Recordset")
		SQLnomat = "SELECT * FROM TB_Materia Where CO_Materia='"&co_materia&"'"
		RSnomat.Open SQLnomat, CON0

no_materia=RSnomat("NO_Materia")		
		%>
                <tr> 
                  <td height="10" colspan="3"><div align="left"><font class="form_dado_texto"> 
                      <%
If co_professor="" or isnull(co_professor) then					
else					  
 		Set RSpro = Server.CreateObject("ADODB.Recordset")
		SQLpro = "SELECT * FROM TB_Professor Where CO_Professor="&co_professor
		RSpro.Open SQLpro, CONp
no_prof=RSpro("NO_Professor")
response.Write(no_prof)
end if		
%>
                      </font></div></td>
                  <td height="10"><div align="left"><font class="form_dado_texto"> 
                      <%response.Write(no_materia)%>
                      </font></div></td>
                  <td height="10" colspan="3"> <div align="left"><font class="form_dado_texto"> 
                      
                      </font></div></td>
                  <td height="10"><font class="form_dado_texto"> 
                    <%

response.Write(aula)		
%>
                    </font></td>
                  <td height="10">&nbsp;</td>
                </tr>
                <tr class="tb_subtit"> 
                  <td height="10" colspan="3">Observa&ccedil;&atilde;o</td>
                  <td height="10">&nbsp;</td>
                  <td height="10" colspan="3">&nbsp;</td>
                  <td height="10">&nbsp;</td>
                  <td height="10">&nbsp;</td>
                </tr>
                <tr > 
                  <td height="10" colspan="9"><div align="center"><font class="form_dado_texto">
                      <%response.Write(observa)		
%>
                      </font></div></td>
                </tr>
              </table></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2"><hr></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2"><div align="center"> 
                <table width="1000" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
                        <tr> 
                          <td width="33%"> <div align="center"> 
                              <input name="Submit" type="submit" class="botao_cancelar"value="Voltar">
                            </div></td>
                          <td width="34%">
<div align="center"></div></td>
                          <td width="33%"><div align="center"></div></td>
                        </tr>
                      </table>
</td>
                  </tr>
                </table>
                <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> </font></div></td>
          </tr>
        </table>
        
      </td>
    </tr>
</form>
<%end if%>
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