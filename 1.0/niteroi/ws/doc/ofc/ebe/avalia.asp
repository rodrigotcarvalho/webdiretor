<%'On Error Resume Next%>

<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->
<!--#include file="../../../../inc/funcoes7.asp"-->
<!--#include file="../../../../../global/conta_alunos.asp"-->
<!--#include file="../../../../../global/tabelas_escolas.asp"-->
<!--#include file="../../../../../global/notas_calculos_diversos.asp"-->
<!--#include file="../../../../../global/funcoes_diversas.asp" -->
<%



	opt = request.QueryString("opt")
	obr = request.QueryString("obr")	
	nvg=session("nvg")
	session("nvg")=nvg
	ano_letivo = session("ano_letivo") 	
	nivel=4
	
	dados_obr=split(obr,"$!$")
	
	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set CON0= Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0

	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1

	Set CONt = Server.CreateObject("ADODB.Connection") 
	ABRIRt = "DBQ="& CAMINHO_t & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONt.Open ABRIRt	

	opt=opt*1
	if opt=1 then
		cod_cons = dados_obr(0)
		periodo = dados_obr(1)	
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& cod_cons &" and TB_Matriculas.NU_Ano="&ano_letivo
		RS.Open SQL, CON1	
			
		no_aluno= RS("NO_Aluno")
		unidade= RS("NU_Unidade")
		curso= RS("CO_Curso")
		co_etapa= RS("CO_Etapa")
		turma= RS("CO_Turma")	
		
		obr_cons="a$!$"&cod_cons	
		
		SESSION("aluno_boletim")=cod_cons
	else	
		unidade = dados_obr(0)
		curso = dados_obr(1)
		co_etapa = dados_obr(2)
		turma = dados_obr(3)
		periodo = dados_obr(4)	
		obr_cons="t"

		if isnull(turma) or turma="" then
			response.Redirect("gera_base.asp?opt=rgnrt&obr="&obr_cons&"$$$"&unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma&"$!$"&periodo)		
		end if
		
		SESSION("aluno_boletim")=NULL	
	end if
		obr_mapa=unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma&"$!$"&periodo&"$!$"&ano_letivo&"$!$boletim"			
	
		obr_ask=obr_cons&"$$$"&obr_mapa
			
		
	call navegacao (CON,nvg,nivel)
	navega=Session("caminho")		
	
	no_unidade= GeraNomes("U",unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)
	no_curso=GeraNomes("C",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro) 	
	no_etapa=GeraNomes("E",curso,co_etapa,variavel3,variavel4,variavel5,CON0,outro) 		


	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Boletim_Cabecalho where NU_Unidade="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'"	
	RS.Open SQL, CONt

		obr_err=obr_cons&"$$$"&no_unidade&"$!$"&no_curso&"$!$"&no_etapa&"$!$"&turma&"$!$0$$$"&obr_mapa		
		ano_liberado=verifica_ano_letivo(ano_letivo,"&nbsp;","&nbsp;","&nbsp;","&nbsp;",CON,"con", "&nbsp;")
	
	if RS.eof then
		msg="rgnrt"

		if ano_liberado="L" then
			response.Redirect("gera_base.asp?opt=rgnrt&obr="&obr_cons&"$$$"&obr_mapa)		
		elseif ano_liberado="B" then

			response.Redirect("index.asp?nvg=WS-DO-DPA-EBE&opt=err707&obr="&obr_err)		
		else
			response.Redirect("index.asp?nvg=WS-DO-DPA-EBE&opt=err9713&obr="&obr_err)
		end if

	else
		msg="cons"
		data_grav=RS("DA_Grav")
		hora_grav=RS("HO_Grav")	
		
		if opt=01 then
			dados_msg=obr_cons&"$$$"&no_aluno&"#!#"&cod_cons&"#!#"&data_grav&"#!#"&hora_grav&"$$$"&obr_mapa				
		else
			dados_msg=obr_cons&"$$$"&no_unidade&"#!#"&no_curso&"#!#"&no_etapa&"#!#"&data_grav&"#!#"&hora_grav&"$$$"&obr_mapa		
		end if

		session ("dados_msg")=dados_msg
		onLoad="onLoad=""redimensiona()"""	
		
		if ano_liberado="L" then
			response.Redirect("index.asp?nvg=WS-DO-DPA-EBE&opt=ask&obr="&obr_ask)		
		elseif ano_liberado="B" then
			response.Redirect("index.asp?nvg=WS-DO-DPA-EBE&opt=acc&obr="&obr_ask&"$!$"&ano_letivo)
		else
			response.Redirect("index.asp?nvg=WS-DO-DPA-EBE&opt=err9713&obr="&obr_err)
		end if			
	end if	



%>	
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">	
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript">
function MM_openBrWindow(theURL,winName) { //v2.0
var largura=screen.availWidth-10
var altura=screen.availHeight-70
url=theURL+"$!$"+largura+"$!$"+altura
  window.open(url,winName,'status=yes,scrollbars=yes,resizable=yes,width='+largura+',height='+altura+',top=0,left=0,bReplace=true');

	go_there();
}
function waitPreloadPage() { //DOM
	if (document.getElementById){
	document.getElementById('prepage').style.visibility='hidden';
	}else{
		if (document.layers){ //NS4
		document.prepage.visibility = 'hidden';
		}
		else { //IE4
		document.all.prepage.style.visibility = 'hidden';
		}
	}
}

function go_there()
{
// var where_to= confirm("<%'response.Write(javascript)%>");
// if (where_to== true)
// {
   window.location="index.asp?nvg=WA-PF-CN-ACC&opt=acc&obr=<%response.Write(obr_mapa)%>";
// }
// else
// {
//  window.location="<%'response.Write("avalia.asp?opt=rgnrt")%>";
//  }
}
function MM_showHideLayers() { //v9.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) 
  with (document) if (getElementById && ((obj=getElementById(args[i]))!=null)) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
    obj.visibility=v; }
}
function redimensiona(){
//o 120 e se refere ao tamanho de cabeçalho do navegador
    y = parseInt((screen.availHeight - 120 - 135 - 70 - 40));
    document.getElementById('carregando_fundo').style.height = y;
}
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" <%response.Write(onload)%>>
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
  <tr>                   
    <td height="10"> 
      <%
	if msg="cons" then 	  
	  	call mensagens(nivel,654,0,dados_msg)	
	else
	  	call mensagens(nivel,653,0,dados_msg)		
	end if		
	%>
    </td>
                  </tr> 
<tr>

            <td valign="top">
                
              <table width="1000" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td width="653" height="15" class="tb_tit">Grade de Aulas</td>
          </tr>
          <tr> 
            <td height="413" valign="top">
    <div id="carregando"  align="center" style="position:absolute;  top: 200px; width:1000px; z-index: 4; height: 150px; visibility: hidden;">
				  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="75" height="75" vspace="80" title="Carregando">
				    <param name="movie" value="../../../../img/carregando.swf">
				    <param name="quality" value="high">
				    <param name="wmode" value="transparent">
				    <embed src="../../../../img/carregando.swf" width="75" height="75" vspace="80" quality="high" wmode="transparent" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash"></embed>
			      </object>
			 </div>              
				<div id="carregando_fundo" align="center" style="position:absolute; width:1000px; z-index: 3; height: 150px; visibility: hidden; background-color:#FFF; top: 250px;  filter: Alpha(Opacity=90, FinishOpacity=100, Style=0, StartX=0, StartY=100, FinishX=100, FinishY=100);">
			 </div>   
            </td>
          </tr>
        </table>
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
'response.redirect("../../../../inc/erro.asp")
end if
%>
</body>
</html>