<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/parametros.asp"-->
<!--#include file="../../../../inc/utils.asp"-->
<!--#include file="../../../../inc/bd_parametros.asp"-->
<!--#include file="../../../../inc/bd_pauta.asp"-->
<% 

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
ori = request.QueryString("or")
chave=session("chave")
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

trava=session("trava")
unidade= session("unidades")
curso= session("grau")
etapa= session("serie")
turma= session("turma")
co_materia = session("co_materia")
periodo = session("periodo")
co_prof = session("co_prof")
co_usr = session("co_usr")
tb = session("nota")
session("co_materia")=co_materia
session("unidades")=unidade
session("grau")=curso
session("serie")=etapa
session("turma")=turma
session("periodo")=periodo
session("co_prof") = co_prof 
session("nota") = tb

vetorDatas=session("obr")
session("obr")=vetorDatas






		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
			
		Set CON_AL = Server.CreateObject("ADODB.Connection") 
		ABRIR_AL = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_AL.Open ABRIR_AL		

		Set CONg = Server.CreateObject("ADODB.Connection") 
		ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONg.Open ABRIRg		
		
		Set CON0= Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CONp= Server.CreateObject("ADODB.Connection") 
		ABRIRp = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONp.Open ABRIRp		
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")	

'		Set RSper = Server.CreateObject("ADODB.Recordset")
'		SQLper = "SELECT * FROM TB_Periodo where NU_Periodo= "&periodo
'		RSper.Open SQLper, CON0
'
'NO_Periodo= RSper("NO_Periodo")
'dataInicio = RSper("DA_Inicio_Periodo")
'dataFim = RSper("DA_Fim_Periodo")
'
'vetorInicioPeriodo = split(dataInicio,"/")
'diaInicial = vetorInicioPeriodo(0)
'mesInicial = vetorInicioPeriodo(1)
'anoInicial = vetorInicioPeriodo(2)
'
'vetorFimPeriodo = split(dataFim,"/")
'diaFinal = vetorFimPeriodo(0)
'mesFinal = vetorFimPeriodo(1)
'anoFinal = vetorFimPeriodo(2)
'
'if isnull(dataInicio) or dataInicio="" then
'
'else
'	dataInicio = formata(dataInicio,"DD/MM/YYYY")
'end if
'
'if isnull(dataFim) or dataFim="" then
'
'else
'	dataFim = formata(dataFim,"DD/MM/YYYY")
'end if





		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"'"
		Set RS = CONg.Execute(CONEXAO)


if RS.EOF then
response.Write("<div align=center><font size=2 face=Courier New, Courier, mono  color=#990000><b>Esta turma não está disponível no momento</b></font><br")
response.Write("<font size=2 face=Courier New, Courier, mono  color=#990000><a href=javascript:window.history.go(-1)>voltar</a></font></div>")

else
coordenador = RS("CO_Cord")
end if

call navegacao (CON,chave,nivel)
navega=Session("caminho")

'datas_periodo = diasPeriodo(periodo)
'datas_formatado = diasPeriodoFormatado(periodo,", ","DD/MM/YYYY")

call GeraNomes(co_materia,unidade,curso,etapa,CON0)

no_materia= session("no_materia")
no_unidade= session("no_unidades")
no_curso= session("no_grau")
no_etapa= session("no_serie")


nome_prof = session("nome_prof") 
tp=	session("tp")

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("m", now) 
data = dia &"/"& mes &"/"& ano
horario = hora & ":"& min
acesso_prof = session("acesso_prof")



		
		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE CO_Professor= "& co_prof &"AND NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' AND CO_Materia_Principal = '"& co_materia &"'"
		Set RS = CONg.Execute(CONEXAO)


planilha_notas = RS("TP_Nota")

bancoPauta = escolheBancoPauta(planilha_notas,"M",p_outro)
caminhoBancoPauta = verificaCaminhoBancoPauta(bancoPauta,"M",p_outro)

		Set RS8 = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT * FROM TB_Materia where CO_Materia='"& co_materia &"'"
		RS8.Open SQL8, CON0

		if RS8.EOF then
			response.Write(co_materia&" não possui nome cadastrado<br>")				
		else
			co_mat_prin= RS8("CO_Materia_Principal")
		end if
		
		if co_mat_prin ="" or isnull(co_mat_prin) then
			co_mat_prin=co_materia
		end if
session("co_mat_prin")=co_mat_prin		
		'if P_DATA_AULA="" then
'			wrkQtdAulasLancadas = 0
'			qtdAulasForm=1
'			data_Pauta = ""	
'			le_tabelas="N"	
'		else
'
			Set CONPauta = Server.CreateObject("ADODB.Connection") 
			ABRIRPauta = "DBQ="& caminhoBancoPauta & ";Driver={Microsoft Access Driver (*.mdb)}"
			CONPauta.Open ABRIRPauta
'			
'			Set RSP = Server.CreateObject("ADODB.Recordset")
'			SQL = "Select TB_Pauta_Aula.NU_Pauta, TB_Pauta_Aula.DT_Aula from TB_Pauta INNER JOIN TB_Pauta_Aula on TB_Pauta.NU_Pauta=TB_Pauta_Aula.NU_Pauta WHERE DT_Aula = #"&P_DATA_AULA&"# AND CO_Professor  = "& co_prof &" AND CO_Materia_Principal = '"& co_mat_prin &"' AND CO_Materia = '"& co_materia &"' AND NU_Unidade  = "& unidade &" AND CO_Curso  = '"& curso &"' AND CO_Etapa  = '"& etapa &"' AND CO_Turma  = '"& turma &"' AND NU_Periodo = "& periodo		
'			Set RSP = CONPauta.Execute(SQL)

'			if RSP.EOF  then
'				wrkQtdAulasLancadas = 1
'				data_Pauta = ""
'				le_tabelas="N"					
'			else
'				NU_Pauta = RSP("NU_Pauta")
'				data_Pauta = RSP("DT_Aula")
'				
'				Set RSP = Server.CreateObject("ADODB.Recordset")
'				SQL = "Select MAX(TB_Pauta_Aula.NU_Tempo) as TotalTempos from TB_Pauta INNER JOIN TB_Pauta_Aula on TB_Pauta.NU_Pauta=TB_Pauta_Aula.NU_Pauta WHERE DT_Aula = #"&P_DATA_AULA&"# AND CO_Professor  = "& co_prof &" AND CO_Materia_Principal = '"& co_mat_prin &"' AND CO_Materia = '"& co_materia &"' AND NU_Unidade  = "& unidade &" AND CO_Curso  = '"& curso &"' AND CO_Etapa  = '"& etapa &"' AND CO_Turma  = '"& turma &"' AND NU_Periodo = "& periodo		
'				Set RSP = CONPauta.Execute(SQL)				
'				wrkQtdAulasLancadas = RSP("TotalTempos")	
'				le_tabelas="S"	
'			end if
'			qtdAulasForm=wrkQtdAulasLancadas
'		end if	

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
			  
        <form action="bd.asp?opt=e" method="post" name="busca" id="busca">
          <tr>
      <td valign="top"> 
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,9715,0,0) %>
    </td>
			  </tr>
          <tr> 
            <td colspan="2"><table width="100%" border="0" cellspacing="0">
                <tr class="tb_subtit"> 
                  <td width="230" height="10"><div align="center"><strong>PER&Iacute;ODO </strong></div></td>
                  <td width="145"><div align="center"><strong>UNIDADE </strong></div></td>
                  <td width="145"><div align="center"><strong>CURSO </strong></div></td>
                  <td width="145"><div align="center"><strong>ETAPA </strong></div></td>
                  <td width="145"><div align="center"><strong>TURMA </strong></div></td>
                  <td width="190"><div align="center"><strong>DISCIPLINA</strong></div></td>
                </tr>
                <tr class="tb_corpo"> 
                  <td width="230" height="10"><div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%
		Set RSper = Server.CreateObject("ADODB.Recordset")
		SQLper = "SELECT * FROM TB_Periodo where NU_Periodo= "&periodo
		RSper.Open SQLper, CON0

NO_Periodo= RSper("NO_Periodo")
'dataInicio = RSper("DA_Inicio_Periodo")
'dataFim = RSper("DA_Fim_Periodo")
'
'if isnull(dataInicio) or dataInicio="" then
'
'else
'	dataInicio = formata(dataInicio,"DD/MM/YYYY")
'end if
'
'if isnull(dataFim) or dataFim="" then
'
'else
'	dataFim = formata(dataFim,"DD/MM/YYYY")
'end if

response.Write(NO_Periodo)%>
                  </font></div></td>
                  <td width="145"><div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%response.Write(no_unidade)%>
                  </font></div></td>
                  <td width="145"><div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%response.Write(no_curso)%>
                  </font></div></td>
                  <td width="145"><div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%
response.Write(no_etapa)%>
                  </font></div></td>
                  <td width="145"><div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%
response.Write(turma)%>
                  </font></div></td>
                  <td width="190"><div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <%

response.Write(no_materia)%>
                  </font></div></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td width="653" bgcolor="#FFFFFF">&nbsp;</td>
            <td width="113">&nbsp;</td>
          </tr>
          <tr> 
            <td height="10" colspan="2" class="tb_tit"
>Conte&uacute;dos   a serem exclu&iacute;dos</td>
          </tr>
          <tr> 
            <td colspan="2"><table width="1000" border="0" cellspacing="0" cellpadding="0">
                <tr class="tb_subtit"> 
                  <td width="30" height="10"><div align="center"></div></td>
                  <td width="180"> <div align="center">Data</div></td>
                  <td width="350"> <div align="center">Conte&uacute;do</div></td>
                </tr>
                <%

check = 2				
vetorExclui = split(vetorDatas,", ")
conta_ocorr=0
for i =0 to ubound(vetorExclui)

dataExibe = replace(vetorExclui(i),".","/")
vetorDataExibe = split(dataExibe,"/")

dataConsulta = vetorDataExibe(1)&"/"&vetorDataExibe(0)&"/"&vetorDataExibe(2)

		if check mod 2 =0 then
			classe = "tb_fundo_linha_par" 
		else 
			classe ="tb_fundo_linha_impar"
		end if 

		Set RSp = Server.CreateObject("ADODB.Recordset")
		SQLp = "SELECT * FROM TB_Professor WHERE CO_Professor ="& co_prof
		RSp.Open SQLp, CONp
		
	IF RSp.EOF then
		no_professor=""
	else
		co_professor=RSp("CO_Usuario")



		Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& co_professor
		RSu.Open SQLu, CON

		IF RSu.EOF then
			no_professor=""
		else
			no_professor=RSu("NO_Usuario")
		end if	
	end if		
	
	
	Set RSP = Server.CreateObject("ADODB.Recordset")
	SQLP = "Select TX_Aula from TB_Materia_Lecionada WHERE DT_Aula = #"&dataConsulta&"# AND CO_Professor  = "& co_prof &" AND CO_Materia_Principal = '"& co_mat_prin &"' AND CO_Materia = '"& co_materia &"' AND NU_Unidade  = "& unidade &" AND CO_Curso  = '"& curso &"' AND CO_Etapa  = '"& etapa &"' AND CO_Turma  = '"& turma &"' AND NU_Periodo = "& periodo

	Set RSP = CONPauta.Execute(SQLP)
	
	While NOT RSP.EOF 	
			tx_aula = RSP("TX_Aula")	
%>
                <tr class="<%=classe%>"> 
                  <td width="30" align="center">&nbsp;</td>
                  <td width="180" align="center"> <div align="center"> 
                      <%response.Write(dataExibe)%>
                    </div></td>
                  <td width="350" align="center"><div align="center"> 
                      <%response.Write(left(tx_aula,200))%>
                    </div></td>
                </tr>
<%

	check = check+1
	RSP.MOVENEXT
	WEND
Next	
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
                        <input name="alterar" type="submit" class="botao_cancelar" id="alterar" onClick="MM_goToURL('parent','notas.asp?opt=vt');return document.MM_returnValue" value="Voltar">
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