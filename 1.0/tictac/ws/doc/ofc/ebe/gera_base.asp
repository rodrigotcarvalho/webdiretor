<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 900 'valor em segundos
%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/parametros.asp"-->
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
	verifica_periodos="s"
	
aluno_temp=SESSION("aluno_boletim")
SESSION("aluno_boletim")=aluno_temp		

	if opt="rgnrt" then	
		separa_dados=split(obr,"$$$")
		tipo_busca=separa_dados(0)	
		dados_funcao=split(separa_dados(1),"$!$")
	
		unidade = dados_funcao(0)
		curso = dados_funcao(1)
		co_etapa = dados_funcao(2)
		turma = dados_funcao(3)
		periodo_form = dados_funcao(4)
		
		if periodo_form=0 then
			verifica_periodos="n"
		end if
	end if

	obr_mapa=tipo_busca&"$$$"&unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma&"$!$"&periodo_form&"$!$"&ano_letivo	
	
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
	
'	Set RS = Server.CreateObject("ADODB.Recordset")
'	SQL = "SELECT * FROM TB_Boletim_Cabecalho where NU_Unidade="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'"
'	RS.Open SQL, CONt
'	
'	if RS.eof or opt="rgnrt" then
	if (co_etapa="999990" or co_etapa="" or isnull(co_etapa)) or (curso="999990" or curso="" or isnull(curso)) then	

	else
		tp_divisao_ano=tipo_divisao_ano(curso,co_etapa,"tp_modelo")

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Periodo where TP_Modelo ='"&tp_divisao_ano&"' ORDER BY NU_Periodo"
		RS0.Open SQL0, CON0
		check_periodo=1
		WHILE NOT RS0.EOF
			periodo=RS0("NU_Periodo")
			
			if verifica_periodos="n" then
				if check_periodo=1 then
					vetor_periodo=periodo
				else
					vetor_periodo=vetor_periodo&"#!#"&periodo
				end if
			else	
				if check_periodo=1 then
					vetor_periodo=periodo
				else
					periodo=periodo*1
					periodo_form=periodo_form*1
					if periodo>periodo_form then
						vetor_periodo=vetor_periodo
					else
						vetor_periodo=vetor_periodo&"#!#"&periodo
					end if	
				end if			
			end if
			check_periodo=check_periodo+1 
		RS0.MOVENEXT
		WEND				
	end if

	if opt="rgnrt" then

		msg="rgnrt"
		
		if turma="999990" or turma="" or isnull(turma) then
			if co_etapa="999990" or co_etapa="" or isnull(co_etapa) then
				if curso="999990" or curso="" or isnull(curso) then		
					if unidade="999990" or unidade="" or isnull(unidade) then
						response.redirect(origem&"index.asp?nvg="&nvg&"&opt=err2")
					else	
					
					
						Set RS0 = Server.CreateObject("ADODB.Recordset")
						SQL0 = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&" ORDER BY CO_Curso,CO_Etapa"
						RS0.Open SQL0, CON0
						check_motriz=1
						WHILE NOT RS0.EOF
							curso=RS0("CO_Curso")
							co_etapa=RS0("CO_Etapa")
							
							Set RS0t = Server.CreateObject("ADODB.Recordset")
							SQL0t = "SELECT * FROM TB_Turma where NU_Unidade="&unidade&" AND CO_Curso ='"&curso&"' AND CO_Etapa ='"&co_etapa&"' ORDER BY CO_Turma"
							RS0t.Open SQL0t, CON0							
							WHILE NOT RS0t.EOF								
								turma=RS0t("CO_Turma")	
	
								if check_motriz=1 then
									vetor_motriz=unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
								else
									vetor_motriz=vetor_motriz&"#$#"&unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
								end if
								check_motriz=check_motriz+1 
							RS0t.MOVENEXT
							WEND	
						RS0.MOVENEXT
						WEND					
						RS0.Close
						Set RS0 = Nothing	
					end if		
				else	
					Set RS0 = Server.CreateObject("ADODB.Recordset")
					SQL0 = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&" AND CO_Curso ='"&curso&"' ORDER BY CO_Etapa"
					RS0.Open SQL0, CON0
					check_motriz=1
					WHILE NOT RS0.EOF
						co_etapa=RS0("CO_Etapa")					
						Set RS0t = Server.CreateObject("ADODB.Recordset")
						SQL0t = "SELECT * FROM TB_Turma where NU_Unidade="&unidade&" AND CO_Curso ='"&curso&"' AND CO_Etapa ='"&co_etapa&"' ORDER BY CO_Turma"
						RS0t.Open SQL0t, CON0							
						WHILE NOT RS0t.EOF								
							turma=RS0t("CO_Turma")	
	
							if check_motriz=1 then
								vetor_motriz=unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
							else
								vetor_motriz=vetor_motriz&"#$#"&unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
							end if
							check_motriz=check_motriz+1 
						RS0t.MOVENEXT
						WEND	
					RS0.MOVENEXT
					WEND
					
					RS0.Close
					Set RS0 = Nothing					
				end if						
			else	
		
				Set RS0t = Server.CreateObject("ADODB.Recordset")
				SQL0t = "SELECT * FROM TB_Turma where NU_Unidade="&unidade&" AND CO_Curso ='"&curso&"' AND CO_Etapa ='"&co_etapa&"' ORDER BY CO_Turma"
				RS0t.Open SQL0t, CON0					
						
				check_motriz=1			
				
				WHILE NOT RS0t.EOF								
					turma=RS0t("CO_Turma")	
	
					if check_motriz=1 then
						vetor_motriz=unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
					else
						vetor_motriz=vetor_motriz&"#$#"&unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
					end if
					check_motriz=check_motriz+1 
				
				RS0t.MOVENEXT
				WEND	
			end if	
			RS0t.Close
			Set RS0t = Nothing	
		ELSE
			vetor_motriz=unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma				
		end if					
'response.Write(vetor_motriz&"-"&vetor_periodo)
'response.End()
		conjunto_dados=split(vetor_motriz,"#$#")
		
		for i=0 to ubound(conjunto_dados)	
			dados_select=split(conjunto_dados(i),"#!#")
			unidade=dados_select(0)
			curso=dados_select(1)
			co_etapa=dados_select(2)
			turma=dados_select(3)		
	
			cria_dados=grava_ficha(unidade, curso, co_etapa, turma, vetor_periodo)
			
			if cria_dados<>"ok" then
				response.Write("avalia.asp?ln100 - ERRO na cria��o dos dados!")
				response.end()
			end if	
		NEXT	
'response.End()
		if cria_dados="ok" then
			if ubound(conjunto_dados)> 0 then
				resultado="acc_mult"			
			else
				resultado="acc"
			end if
			response.redirect("index.asp?nvg=WS-DO-DPA-EBE&opt="&resultado&"&obr="&obr_mapa)
	'		response.Write("OK")
		else
			response.Write("avalia.asp?ln200 - ERRO na cria��o dos dados!")
			response.end()
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
   window.location="index.asp?nvg=WS-DO-DPA-EBE&opt=acc&obr=<%response.Write(obr_mapa)%>";
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
	  if msg="rgnrt" then
	  	call mensagens(nivel,653,1,0) 	  
	  elseif msg="nrgnrt" then
	  	call mensagens(nivel,9712,2,dados_msg) 	
	  else
	  	call mensagens(nivel,652,0,dados_msg)	
	  end if
	%>
    </td>
                  </tr> 
<tr>

            <td valign="top">
                
              <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
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