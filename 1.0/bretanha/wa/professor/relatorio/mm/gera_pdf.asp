<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->
<!--#include file="../../../../inc/funcoes7.asp"-->
<!--#include file="../../../../../global/conta_alunos.asp"-->
<!--#include file="../../../../../global/tabelas_escolas.asp"-->
<!--#include file="../../../../../global/notas_calculos_diversos.asp"-->
<%
	obr = request.QueryString("obr")	
	nvg=session("nvg")
	session("nvg")=nvg
	nivel=4

	dados_funcao=split(obr,"$!$")

	unidade = dados_funcao(0)
	curso = dados_funcao(1)
	co_etapa = dados_funcao(2)
	turma = dados_funcao(3)
	periodo = dados_funcao(4)
	acumulado = dados_funcao(5)
	qto_falta = dados_funcao(6)	
	ano_letivo = dados_funcao(7)
	
	obr_mapa=unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma&"$!$"&periodo&"$!$"&acumulado&"$!$"&qto_falta&"$!$"&ano_letivo	
	obr_log=unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&periodo&"_"&ano_letivo

	call GravaLog (nvg,obr_log)	

response.redirect("../../../../relatorios/swd102.asp?obr="&obr_mapa)
%>	
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">	
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
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
	  	call mensagens(nivel,647,1,0) 	  
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