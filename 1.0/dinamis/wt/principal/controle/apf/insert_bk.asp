<%' Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes_comuns.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%
opt = request.QueryString("opt")
nvg=request.QueryString("nvg")
session("nvg")=nvg

ano_letivo = session("ano_letivo")
session("ano_letivo") = ano_letivo

Server.ScriptTimeout = 900 'valor em segundos

nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

nivel = 4

ano_info=nivel&"-"&nvg&"-"&ano_letivo



		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0



call navegacao (CON,nvg,nivel)
navega=Session("caminho")	

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="file:../../../../img/mm_menu.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
<script type="text/javascript" src="../../../../js/global.js"></script>
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
function submitforminterno()  
{
   var f=document.forms[3]; 
      f.submit(); 
	  
}
//-->
</script>
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
//-->
</script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../../img/menu_r1_c2_f3.gif','../../../../img/menu_r1_c2_f2.gif','../../../../img/menu_r1_c2_f4.gif','../../../../img/menu_r1_c4_f3.gif','../../../../img/menu_r1_c4_f2.gif','../../../../img/menu_r1_c4_f4.gif','../../../../img/menu_r1_c6_f3.gif','../../../../img/menu_r1_c6_f2.gif','../../../../img/menu_r1_c6_f4.gif','../../../../img/menu_r1_c8_f3.gif','../../../../img/menu_r1_c8_f2.gif','../../../../img/menu_r1_c8_f4.gif','../../../../img/menu_direita_r2_c1_f3.gif','../../../../img/menu_direita_r2_c1_f2.gif','../../../../img/menu_direita_r2_c1_f4.gif','../../../../img/menu_direita_r4_c1_f3.gif','../../../../img/menu_direita_r4_c1_f2.gif','../../../../img/menu_direita_r4_c1_f4.gif','../../../../img/menu_direita_r6_c1_f3.gif','../../../../img/menu_direita_r6_c1_f2.gif','../../../../img/menu_direita_r6_c1_f4.gif')">
<% call cabecalho (nivel)
	  %>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr>
    <td height="10" class="tb_caminho"><font class="style-caminho">
      <%
	  response.Write(navega)

%>
      </font></td>
  </tr>
  <%if opt="a1" then%>
  <tr>
    <td height="10"><%	call mensagens(4,903,2,0) %></td>
  </tr>
  <%elseif opt="a2" then%>
  <tr>
    <td height="10"><%	call mensagens(4,904,2,0) %></td>
  </tr>
  <%end if%>
  </tr>
  <tr>
    <td valign="top">&nbsp;
</td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
</body>
</html>
<%
mes = DatePart("m", now) 
dia = DatePart("d", now) 

dia=dia*1
mes=mes*1
%>
<%
'parametros'-----------------------------------------------
'caminho 

PastaDestino = CAMINHO_tp
if opt="a1" then
	str_arquivo= PastaDestino&"POSICAOWEB.txt" 
	caminho_insert=CAMINHO_pf

else
	arquivo_p_inserido = request.QueryString("aep")
	str_arquivo= PastaDestino&"BOLETOWEB.txt"
	caminho_insert=CAMINHO_bl

end if	

'delimitador
str_delimitador = ";" 

'Função de Ler o arquivo texto e fazer insert na tabela do access
'conexao'-----------------------------------------------
Set bd = Server.CreateObject("ADODB.Connection")
localbd = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source ="&caminho_insert
bd.Open localbd 

if err.number <> 0 then        
	Response.Write "erro na conexao"
end if


conta_registros=0

 'Cria o objeto
Set FSO = Server.CreateObject("Scripting.FileSystemObject")
'Verificando e inserindo'-----------------------------------------------
If fso.FileExists(str_arquivo) Then                
	Set TXT = FSO.OpenTextFile(str_Arquivo)                
	while not txt.AtEndOfStream
		campo =  txt.readLine                
'		Response.write campo & chr(13)& "<br>"                
		ar_camp = split(campo,str_delimitador)  
   		
		if ubound(ar_camp) > 0 then 

			if opt="a1" then
				url="insert.asp?nvg="&nvg&"&aep=s&opt=a2"              
				CO_Matricula_Escola = ar_camp(0)                        
				DA_Vencimento = ar_camp(1)                        
				NU_Cota = ar_camp(2)                        
				NO_Lancamento = ar_camp(3)     
				VA_Compromisso = ar_camp(4)                        
				DA_Realizado = ar_camp(5)                        
				VA_Realizado = ar_camp(6)                        
				Mes = ar_camp(7)     
				
				Set RSc = Server.CreateObject("ADODB.Recordset")		  		                   
				sqlc = "Select * From TB_Posicao"      
				RSc.Open sqlc, bd			                  
				
				if RSc.EOF then 
					Set RS = server.createobject("adodb.recordset")		
					RS.open "TB_Posicao", bd, 2, 2 'which table do you want open
					RS.addnew
						RS("CO_Matricula_Escola") = CO_Matricula_Escola
						RS("DA_Vencimento") = DA_Vencimento
						RS("NU_Cota") = NU_Cota
						RS("NO_Lancamento") = NO_Lancamento
						RS("VA_Compromisso")=VA_Compromisso				
						RS("DA_Realizado")=DA_Realizado
						RS("VA_Realizado")=VA_Realizado
						RS("Mes")=Mes			
					RS.update
					set RS=nothing	
				else	
					if conta_registros=0 then	
						Set RSd = Server.CreateObject("ADODB.Recordset")
						sqld = "DELETE * from TB_Posicao"  
						Set RS0 = bd.Execute(sqld)		
					end if
					
					Set RS = server.createobject("adodb.recordset")		
					RS.open "TB_Posicao", bd, 2, 2 'which table do you want open
					RS.addnew
'					response.Write("'"&CO_Matricula_Escola&"'<BR>")		
'					response.Write("'"&DA_Vencimento&"'<BR>")							
'					response.Write("'"&DA_Realizado&"'<BR>")
					if isnull(DA_Realizado) or DA_Realizado="" then
						DA_Realizado=NULL
					end if	
					
						RS("CO_Matricula_Escola") = CO_Matricula_Escola
						RS("DA_Vencimento") = DA_Vencimento
						RS("NU_Cota") = NU_Cota
						RS("NO_Lancamento") = NO_Lancamento
						RS("VA_Compromisso")=VA_Compromisso				
						RS("DA_Realizado")=DA_Realizado
						RS("VA_Realizado")=VA_Realizado
						RS("Mes")=Mes			
					RS.update
					set RS=nothing				
				
				end if	
			else	
				url="index.asp?nvg="&nvg&"&opt=ok&aep="&arquivo_p_inserido&"&aeb=s"  
				
				if ar_camp(4) ="" or isnull(ar_camp(4)) then
					wrk_bloqueto=NULL	
				else
					wrk_bloqueto=ar_camp(4) 
				end if	
				if ar_camp(6) ="" or isnull(ar_camp(6)) then
					wrk_va_inicial=NULL	
				else
					wrk_va_inicial=ar_camp(6) 
				end if		
				if ar_camp(18) ="" or isnull(ar_camp(18)) then
					wrk_nu_endereco=NULL	
				else
					wrk_nu_endereco=ar_camp(18) 
				end if	
				if ar_camp(23) ="" or isnull(ar_camp(23)) then
					wrk_cep_endereco=NULL	
				else
					wrk_cep_endereco=ar_camp(23) 
				end if		
																		
						
				CO_Matricula_Escola = ar_camp(0)                        
				DA_Vencimento = ar_camp(1)                        
				CO_Turma = ar_camp(2)                        
				NO_Aluno = ar_camp(3)     
				NU_Cota = ar_camp(4)                        
				NU_Bloqueto = wrk_bloqueto                       
				VA_Inicial = wrk_va_inicial                       
				CO_Superior = ar_camp(7)     
				NO_Cedente = ar_camp(8)                        
				CO_Agencia = ar_camp(9)                        
				CO_Conta = ar_camp(10)                        
				DA_Processamento = ar_camp(11)     
				TX_Msg_01 = ar_camp(12)                        
				TX_Msg_02 = ar_camp(13)                        
				TX_Msg_03 = ar_camp(14) 
				TX_Msg_04 = ar_camp(15)                       
				NO_Responsavel = ar_camp(16)   		
				NO_Logradouro_Empresa = ar_camp(17)                   
				NU_Logradouro_Empresa = wrk_nu_endereco                      
				TX_Complemento_Logradouro_Empresa = ar_camp(19)                        
				NO_Bairro_Empresa = ar_camp(20)     
				NO_Cidade_Empresa = ar_camp(21)                        
				SG_UF_Empresa = ar_camp(22)                        
				CO_CEP_Empresa = wrk_cep_endereco                       
				NO_Grau = ar_camp(24)     
				NO_Serie = ar_camp(25)                        
				CO_Barras = ar_camp(26)                        
				CO_CPF = ar_camp(27)                        
				CO_Nosso_Numero = ar_camp(28)   
				TX_Msg_Extra = ar_camp(29)     					  	
'					response.Write("'"&CO_Matricula_Escola&"'<BR>")		
'					response.Write("'"&DA_Vencimento&"'<BR>")							
'					response.Write("'"&CO_CPF&"'<BR>")
					
				Set RSc = Server.CreateObject("ADODB.Recordset")		  		                   
				sqlc = "Select * From TB_Bloqueto"
				RSc.Open sqlc, bd	
				
				if RSc.EOF then 
					Set RSb = server.createobject("adodb.recordset")		
					RSb.open "TB_Bloqueto", bd, 2, 2 'which table do you want open
					RSb.addnew
					
						RSb("CO_Matricula_Escola") = CO_Matricula_Escola
						RSb("DA_Vencimento") = DA_Vencimento
						RSb("CO_Turma") = CO_Turma
						RSb("NO_Aluno") = NO_Aluno
						RSb("NU_Cota")=NU_Cota				
						RSb("NU_Bloqueto")=NU_Bloqueto
						RSb("VA_Inicial")=VA_Inicial
						RSb("CO_Superior")=CO_Superior	 
						RSb("NO_Cedente") = NO_Cedente
						RSb("CO_Agencia") = CO_Agencia
						RSb("CO_Conta")=CO_Conta				
						RSb("DA_Processamento")=DA_Processamento
						RSb("TX_Msg_01")=TX_Msg_01
						RSb("TX_Msg_02")=TX_Msg_02		
						RSb("TX_Msg_03") = TX_Msg_03
						RSb("TX_Msg_04") = TX_Msg_04									
						RSb("NO_Responsavel") = NO_Responsavel
						RSb("NO_Logradouro_Empresa")=NO_Logradouro_Empresa				
						RSb("NU_Logradouro_Empresa")=NU_Logradouro_Empresa
						RSb("TX_Complemento_Logradouro_Empresa")=TX_Complemento_Logradouro_Empresa
						RSb("NO_Bairro_Empresa")=NO_Bairro_Empresa													
						RSb("NO_Cidade_Empresa") = NO_Cidade_Empresa
						RSb("SG_UF_Empresa") = SG_UF_Empresa
						RSb("CO_CEP_Empresa")=CO_CEP_Empresa				
						RSb("NO_Grau")=NO_Grau
						RSb("NO_Serie")=NO_Serie
						RSb("CO_Barras")=CO_Barras		
						RSb("CO_CPF")=CO_CPF
						RSb("CO_Nosso_Numero")=CO_Nosso_Numero	
						RSb("TX_Msg_Extra")=TX_Msg_Extra															
					RSb.update
					set RSb=nothing	
				else
					if conta_registros=0 then							
						Set RSd = Server.CreateObject("ADODB.Recordset")
						sqld = "DELETE * from TB_Bloqueto"  
						Set RS0 = bd.Execute(sqld)		
					end if
					Set RSb = server.createobject("adodb.recordset")		
					RSb.open "TB_Bloqueto", bd, 2, 2 'which table do you want open
					RSb.addnew
					
						RSb("CO_Matricula_Escola") = CO_Matricula_Escola
						RSb("DA_Vencimento") = DA_Vencimento
						RSb("CO_Turma") = CO_Turma
						RSb("NO_Aluno") = NO_Aluno
						RSb("NU_Cota")=NU_Cota				
						RSb("NU_Bloqueto")=NU_Bloqueto
						RSb("VA_Inicial")=VA_Inicial
						RSb("CO_Superior")=CO_Superior	 
						RSb("NO_Cedente") = NO_Cedente
						RSb("CO_Agencia") = CO_Agencia
						RSb("CO_Conta")=CO_Conta				
						RSb("DA_Processamento")=DA_Processamento
						RSb("TX_Msg_01")=TX_Msg_01
						RSb("TX_Msg_02")=TX_Msg_02		
						RSb("TX_Msg_03") = TX_Msg_03
						RSb("TX_Msg_04") = TX_Msg_04												
						RSb("NO_Responsavel") = NO_Responsavel
						RSb("NO_Logradouro_Empresa")=NO_Logradouro_Empresa				
						RSb("NU_Logradouro_Empresa")=NU_Logradouro_Empresa
						RSb("TX_Complemento_Logradouro_Empresa")=TX_Complemento_Logradouro_Empresa
						RSb("NO_Bairro_Empresa")=NO_Bairro_Empresa															
						RSb("NO_Cidade_Empresa") = NO_Cidade_Empresa
						RSb("SG_UF_Empresa") = SG_UF_Empresa
						RSb("CO_CEP_Empresa")=CO_CEP_Empresa				
						RSb("NO_Grau")=NO_Grau
						RSb("NO_Serie")=NO_Serie
						RSb("CO_Barras")=CO_Barras		
						RSb("CO_CPF")=CO_CPF
						RSb("CO_Nosso_Numero")=CO_Nosso_Numero		
						RSb("TX_Msg_Extra")=TX_Msg_Extra			
					RSb.update
					set RSb=nothing				
				
				end if	
			end if		                   
		end if		
		Set ar_camp = nothing                        
		sql = ""      
	conta_registros=conta_registros+1	   
	wend        
	if err.number <> 0 then                
		Response.Write "erro ao fazer o insert"                
		Response.End         
	end if         
	Response.Write "campos inseridos"        
	txt.close        
	bd.Close    
	    
	if opt="a1" then	
		CALL GravaLog (nvg,"POSICAOWEB")
	else
		CALL GravaLog (nvg,"BOLETOWEB")	
	end if	
    
	fso.DeleteFile(str_arquivo)		
	
	Set bd = Nothing        
	Set fso = Nothing

Else  
	if opt="a1" then
		url="insert.asp?nvg="&nvg&"&aep=n&opt=a2"   
	else	 
		if arquivo_p_inserido = "n" then   
			url="index.asp?nvg="&nvg&"&opt=err2"
		else
			url="index.asp?nvg="&nvg&"&aep=s&aeb=n&opt=ok"	
		end if	
	end if	 	
	'response.write "Arquivo não encontrado !!!"
End If
'If compacta="OK" then
'	response.write "Arquivo compactado!!!"
'else
'	response.write "Arquivo não encontrado !!!"
'end if
response.Redirect(url)

%>
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