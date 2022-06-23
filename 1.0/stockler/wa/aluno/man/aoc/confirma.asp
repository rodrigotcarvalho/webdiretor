<%On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->


<!--#include file="../../../../inc/caminhos.asp"-->



<% 

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
ori = request.QueryString("or")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

trava=session("trava")
exclui_ocorrencia=request.form("ocorrencia")

obr=session("obr")
session("obr")=obr
'obr=cod&"?"&ordem&"?"&tp_ocor&"?"&data_de&"?"&hora_de&"?"&data_inicio&"?"&data_ate&"?"&hora_ate&"?"&data_fim
dados= split(obr, "?" )
cod= dados(0)




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

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
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
			  
        <form action="bd.asp?opt=exc" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
          <tr>
      <td valign="top"> 
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,311,0,0) %>
    </td>
			  </tr>
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
                    <%response.Write(codigo)%>
                    </font></td>
                  <td width="6%" height="10"> <div align="right"><font class="form_dado_texto"> 
                      Nome: </font></div></td>
                  <td width="66%" height="10"><font class="form_dado_texto"> 
                    <%response.Write(nome_prof)%>
                    <input name="exclui_ocorrencia" type="hidden" class="textInput" id="nome2"  value="<%=exclui_ocorrencia%>" size="75" maxlength="50">
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
call GeraNomes("PORT",unidade,curso,etapa,CON0)
no_unidades = session("no_unidades")
no_grau = session("no_grau")
no_serie = session("no_serie")


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
                      <%response.Write(ano_aluno)
					  %>
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
              </table></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td height="10" colspan="2" class="tb_tit"
>Ocorr&ecirc;ncias a serem exclu&iacute;das</td>
          </tr>
          <tr> 
            <td colspan="2"><table width="1000" border="0" cellspacing="0" cellpadding="0">
                <tr class="tb_subtit"> 
                  <td width="30" height="10"><div align="center"></div></td>
                  <td width="180"> <div align="center">Data / Hora</div></td>
                  <td width="350"> <div align="center">Ocorr&ecirc;ncia</div></td>
                  <td width="200"> <div align="center">Professor</div></td>
                  <td width="110"><div align="center">Disciplina</div></td>
                  <td width="110"><div align="center">Aula</div></td>
                  <td width="200"> <div align="center">Atendido por</div></td>
                </tr>
                <%
'response.Write(">>"&exclui_ocorrencia)				
check = 2				
vertorExclui = split(exclui_ocorrencia,", ")
conta_ocorr=0
for i =0 to ubound(vertorExclui)

exclui = split(vertorExclui(i),"?")

'obr=cod&"?"&da_ocorrencia&"?"&ho_ocorrencia&"?"&co_ocorrencia
cod = exclui(0)
da_ocorrencia= exclui(1)
ho_ocorrencia= exclui(2)
co_ocorrencia= exclui(3)
				
dados_data=split(da_ocorrencia,"/")
dia=dados_data(0)
mes=dados_data(1)
ano=dados_data(2)

dados_hora=split(ho_ocorrencia,":")
h=dados_hora(0)
m=dados_hora(1)


da_ocorrencia_cons=mes&"/"&dia&"/"&ano

h=h*1
m=m*1


ho_ocorrencia_cons=h&":"&m
'response.Write(ho_ocorrencia_cons)				
		Set RSo = Server.CreateObject("ADODB.Recordset")
		SQLo = "SELECT * FROM TB_Ocorrencia_Aluno WHERE CO_Matricula ="& cod&" AND CO_Ocorrencia ="& co_ocorrencia&" AND (DA_Ocorrencia=#"&da_ocorrencia_cons&"# AND mid(HO_Ocorrencia,1,16)=#12/30/1899 "&ho_ocorrencia_cons&"#)" 
'		SQLo = "SELECT * FROM TB_Ocorrencia_Aluno WHERE (((TB_Ocorrencia_Aluno.CO_Matricula)="& cod&") AND ((TB_Ocorrencia_Aluno.DA_Ocorrencia)=#"&da_ocorrencia_cons&"#) AND ((TB_Ocorrencia_Aluno.HO_Ocorrencia)=#12/30/1899 "&ho_ocorrencia&"#))"
'		SQLo = "SELECT * FROM TB_Ocorrencia_Aluno WHERE CO_Matricula ="& cod&" AND DA_Ocorrencia= #"&da_ocorrencia_cons&"# AND HO_Ocorrencia=#12/30/1899 "&ho_ocorrencia&"#"
		RSo.Open SQLo, CON3

  if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if 
  
co_ocorrencia=RSo("CO_Ocorrencia")
da_ocorrencia=RSo("DA_Ocorrencia")
ho_ocorrencia=RSo("HO_Ocorrencia")
ass_ocorrencia=RSo("CO_Assunto")
au_ocorrencia=RSo("NU_Aula")
cp_ocorrencia=RSo("CO_Professor")
di_ocorrencia=RSo("NO_Materia")
ob_ocorrencia=RSo("TX_Observa")
cu_ocorrencia=RSo("CO_Usuario")

dados_hora=split(ho_ocorrencia,":")
h=dados_hora(0)
m=dados_hora(1)


h=h*1
m=m*1

if h<10 then
h="0"&h
end if

if m<10 then
m="0"&m
end if

ho_ocorrencia=h&":"&m



 
 		Set RSto = Server.CreateObject("ADODB.Recordset")
		SQLto = "SELECT * FROM TB_Tipo_Ocorrencia WHERE CO_Ocorrencia ="& co_ocorrencia
		RSto.Open SQLto, CON0
no_ocorrencia=RSto("NO_Ocorrencia")

if cp_ocorrencia="" or isnull(cp_ocorrencia) or cp_ocorrencia="999999" or cp_ocorrencia=999999  then
	no_professor=""
else

		Set RSp = Server.CreateObject("ADODB.Recordset")
		SQLp = "SELECT * FROM TB_Professor WHERE CO_Professor ="& cp_ocorrencia
		RSp.Open SQLp, CONp
		
	IF RSp.EOF then
	no_professor=""
	else
	co_professor=RSp("CO_Usuario")
	end if


		Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& co_professor
		RSu.Open SQLu, CON

	IF RSu.EOF then
	no_professor=""
	else
	no_professor=RSu("NO_Usuario")
	end if		
end if
			
	if cu_ocorrencia="" or isnull(cu_ocorrencia) then
	else
		Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& cu_ocorrencia
		RSu.Open SQLu, CON

		IF RSu.EOF then
		no_atendido=""
		else
no_atendido=RSu("NO_Usuario")
		end if
		
	end if
opt=cod&"?"&da_ocorrencia&"?"&ho_ocorrencia&"?"&co_ocorrencia

%>
                <tr class="<%=cor%>"> 
                  <td width="30">&nbsp;</td>
                  <td width="180"> <div align="center"> 
                      <%response.Write(da_ocorrencia&", "&ho_ocorrencia)%>
                    </div></td>
                  <td width="350"><div align="center"> <A href="ocorrencia.asp?opt=<%=opt%>" class="linkum"> 
                      <%response.Write(no_ocorrencia)%></A> 
                    </div></td>
                  <td width="200"> <div align="center"> 
                      <%response.Write(no_professor)%>
                    </div></td>
                  <td width="110"><div align="center"> 
                      <%response.Write(di_ocorrencia)%>
                    </div></td>
                  <td width="110"><div align="center"> 
                      <%response.Write(au_ocorrencia)%>
                    </div></td>
                  <td width="200"> <div align="center"> 
                      <%response.Write(no_atendido)%>
                    </div></td>
                </tr>
<%check = check+1
next
%>				
              </table></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></div></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td colspan="2"><hr></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2"><div align="center"> 
                <table width="1000" border="0" align="center" cellspacing="0">
                  <tr> 
                    <td width="391"> <div align="center"> 
                        <input name="alterar" type="submit" class="botao_cancelar" id="alterar" onClick="MM_goToURL('parent','resumo.asp?or=2');return document.MM_returnValue" value="Voltar">
                      </div></td>
                    <td width="391">&nbsp;</td>
                    <td width="218"> <div align="left"> 
                        <input name="Submit" type="submit" class="botao_prosseguir" value="Confirmar">
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