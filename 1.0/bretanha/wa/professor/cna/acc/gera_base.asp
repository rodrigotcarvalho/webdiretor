<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 3600 'valor em segundos
%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->
<!--#include file="../../../../inc/funcoes7.asp"-->
<!--#include file="../../../../../global/conta_alunos.asp"-->
<!--#include file="../../../../../global/tabelas_escolas.asp"-->
<!--#include file="../../../../../global/notas_calculos_diversos.asp"-->
<%
	opt = request.QueryString("opt")
	obr = request.QueryString("obr")	
	nvg=session("nvg")
	session("nvg")=nvg
	nivel=4


	if opt="rgnrt" then

		dados=obr
		dados_funcao=split(obr,"$!$")
	
		unidade = dados_funcao(0)
		curso = dados_funcao(1)
		co_etapa = dados_funcao(2)
		turma = dados_funcao(3)
		periodo = dados_funcao(4)
		acumulado = dados_funcao(5)
		qto_falta = dados_funcao(6)	
		ano_letivo = dados_funcao(7)
	end if
	
	obr_mapa=unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma&"$!$"&periodo&"$!$"&acumulado&"$!$"&qto_falta&"$!$"&ano_letivo	
	
	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
		
	Set CON0= Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0

	Set CONt = Server.CreateObject("ADODB.Connection") 
	ABRIRt = "DBQ="& CAMINHO_t & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONt.Open ABRIRt
		
	call navegacao (CON,nvg,nivel)
	navega=Session("caminho")		
	
	call GeraNomes("PORT",unidade,curso,co_etapa,CON0)	
	no_unidade= session("no_unidades")
	no_curso= session("no_grau")
	no_etapa= session("no_serie")	


	
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Mapao_Disciplinas where NU_Unidade="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'"
	RS.Open SQL, CONt
	
	if RS.eof or opt="rgnrt" then
		msg="rgnrt"
			Set RSapr = Server.CreateObject("ADODB.Recordset")
			SQLapr = "Select * from TB_Regras_Aprovacao WHERE CO_Curso = '"& curso &"' AND CO_Etapa='"&co_etapa&"'"
			Set RSapr = CON0.Execute(SQLapr)
			
			if RSapr.EOF then
				ntvml=0
			else
				ntazl= RSapr("NU_Valor_M1")		
				ntvml= RSapr("NU_Valor_M2")
				peso_m2_m1=RSapr("NU_Peso_Media_M2_M1")
				peso_m2_m2=RSapr("NU_Peso_Media_M2_M2")
				peso_m3_m1=RSapr("NU_Peso_Media_M3_M1")
				peso_m3_m2=RSapr("NU_Peso_Media_M3_M2")
				peso_m3_m3=RSapr("NU_Peso_Media_M3_M3")		
			end if

		cria_dados=grava_ACC(unidade, curso, co_etapa, turma, periodo, acumulado, qto_falta, ntazl, ntvml, 999, peso_m2_m1, peso_m2_m2, peso_m3_m1, peso_m3_m2, peso_m3_m3)

		if cria_dados="ok" then
			response.redirect("carregando.asp?opt=nrgnrt&obr="&obr_mapa)
		else
			response.Write("avalia.asp?ln100 - ERRO na cria��o dos dados!")
			response.end()
		end if	
	else
		if opt="nrgnrt" then
			msg="nrgnrt"
			onload="onLoad=""redimensiona();MM_openBrWindow('mapa.asp?obr="&obr_mapa&"','')"""			
		else
			msg="cons"
			data_grav=RS("DA_Grav")
			hora_grav=RS("HO_Grav")	
			dados_msg=no_unidade&"#!#"&no_curso&"#!#"&no_etapa&"#!#"&data_grav&"#!#"&hora_grav&"#$#"&obr_mapa	
			onLoad="onLoad=""redimensiona()"""			
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
//o 120 e se refere ao tamanho de cabe�alho do navegador
    y = parseInt((screen.availHeight - 120 - 135 - 70 - 40));
    document.getElementById('carregando_fundo').style.height = y;
}
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" <%response.Write(onload)%>>
<% call cabecalho (nivel)
	  %>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
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
	  if msg="rgnrt" then
	  	call mensagens(nivel,647,1,0) 	  
	  elseif msg="nrgnrt" then
	  	call mensagens(nivel,9712,2,dados_msg) 	
	  else
	  	call mensagens(nivel,646,0,dados_msg)	
	  end if
	%>
    </td>
                  </tr> 
<tr>

            <td valign="top">
                
              <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td width="653" height="15" class="tb_tit">Grade de Aulas</td>
          </tr>
          <tr> 
            <td height="413" valign="top">
            	<%' if opt="rgnrt" or opt="nrgnrt" then%>
				<div id="carregando" align="center" style="position:absolute; width:1000px; z-index: 4; height: 150px; visibility: hidden;">
				  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="75" height="75" vspace="80" title="Carregando">
				    <param name="movie" value="../../../../img/carregando.swf">
				    <param name="quality" value="high">
				    <param name="wmode" value="transparent">
				    <embed src="../../../../img/carregando.swf" width="75" height="75" vspace="80" quality="high" wmode="transparent" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash"></embed>
			      </object>
			 </div>              
				<div id="carregando_fundo" align="center" style="position:absolute; width:1000px; z-index: 3; height: 150px; visibility: hidden; background-color:#FFF; filter: Alpha(Opacity=90, FinishOpacity=100, Style=0, StartX=0, StartY=0, FinishX=100, FinishY=100); ">
	   </div>  
                 <%'end if%> 
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