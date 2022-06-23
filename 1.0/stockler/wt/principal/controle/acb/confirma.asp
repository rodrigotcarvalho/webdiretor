<%	'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<% 
nivel=4
concatena_contrato = request.QueryString("cc")
vetor_concatena = split(concatena_contrato,", ")

FOR vc = 0 to ubound(vetor_concatena)
	vetor_temp = split(vetor_concatena(vc),"-")
	vetor_contratos = split(vetor_temp(1),"$")
	contratos = vetor_contratos(1)
	if vc= 0 then
		sql_contratos = contratos
	else
		sql_contratos = sql_contratos&", "&contratos	
	end if	
NEXT	



permissao = session("permissao") 
ano_letivo_wf = session("ano_letivo_wf")
sistema_local=session("sistema_local")
nvg = session("chave")
chave=nvg
session("chave")=chave
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)


cod_form=Session("cod_form")
nome_form=Session("nome_form")
contrato=Session("contrato")
ativos = Session("ativos")
cancelados = Session("cancelados")		
sem_parcelas = Session("sem_parcelas")
so_bolsistas = Session("so_bolsistas")
dia_de= Session("dia_de")
mes_de= Session("mes_de")
ano_de=Session("ano_de")
dia_ate=Session("dia_ate")
mes_ate=Session("mes_ate")
ano_ate=Session("ano_ate")
bolsa=Session("bolsa")
desconto_de=Session("desconto_de")
desconto_ate=session("desconto_ate")
unidade=Session("unidade")
curso=Session("curso")
etapa=Session("etapa")
turma=session("turma")


Session("cod_form")=cod_form
Session("nome_form")=nome_form
Session("contrato")=contrato
Session("ativos")=ativos
Session("cancelados")=cancelados	
Session("sem_parcelas")=sem_parcelas
Session("so_bolsistas")=so_bolsistas
Session("dia_de")=dia_de
Session("mes_de")=mes_de
Session("ano_de")=ano_de
Session("dia_ate")=dia_ate
Session("mes_ate")=mes_ate
Session("ano_ate")=ano_ate
Session("bolsa")=bolsa
Session("desconto_de")=desconto_de
session("desconto_ate") =desconto_ate
Session("unidade")=unidade
Session("curso")=curso
Session("etapa")=etapa
session("turma") =turma


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")		


		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		


		Set CON5 = Server.CreateObject("ADODB.Connection") 
		ABRIR5 = "DBQ="& CAMINHO_cr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON5.Open ABRIR5

		Set CONa = Server.CreateObject("ADODB.Connection") 
		ABRIRa = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONa.Open ABRIRa		

%>
<html>
<head>
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../../../js/mm_menu.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--

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
<form action="bd.asp?opt=e" method="post" name="busca" id="busca">
<%call cabecalho_novo(nivel)
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
      <td valign="top"> 
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
          <tr> 
            <td width="766" height="10" colspan="4" valign="top"> 
              <%call mensagens(nivel,801,0,0) %>
            </td>
          </tr>
          <tr> 
            <td height="10" class="tb_tit"
>Contratos a serem cancelados</td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="tb_subtit">
                <td width="20" align="center"><div align="center"><input name="excluir_contratos" type="hidden" id="excluir_contratos" value="<%response.Write(concatena_contrato)%>"></div></td>
                <td width="100" align="center">Data</td>
                <td width="100" align="center">N&uacute;mero</td>
                <td width="100" align="center">Matr&iacute;cula</td>
                <td width="420" align="center">Aluno</td>
                <td width="50" align="center">Bolsista?</td>
                <td width="50" align="center">Un</td>
                <td width="50" align="center">Curso</td>
                <td width="50" align="center">Etapa</td>
                <td width="50" align="center">Turma</td>
                <td width="50" align="center">Situa&ccedil;&atilde;o</td>
              </tr>
              <%
		Set RSC= Server.CreateObject("ADODB.Recordset")
		SQLC = "SELECT * FROM TB_Contrato where NU_Contrato IN ("&sql_contratos&") order by NU_Ano_Letivo Desc,NU_Contrato"		
		RSC.Open SQLC, CON5
		
		if RSC.EOF then

		else
	
	check=2
	While Not RSC.EoF
	
	 if check mod 2 =0 then
		cor = "tb_fundo_linha_par" 
	 else 
		cor ="tb_fundo_linha_impar"
	 end if
	
		data_contrato = RSC("DT_Contrato")
		ano_contrato = RSC("NU_Ano_Letivo")
		nu_contrato = RSC("NU_Contrato")
		if nu_contrato<100000 then
			if nu_contrato<10000 then
				if nu_contrato<1000 then
					if nu_contrato<100 then
						if nu_contrato<10 then
							nu_contrato="00000"&nu_contrato							
						else
							nu_contrato="0000"&nu_contrato					
						end if						
					else
						nu_contrato="000"&nu_contrato					
					end if	
				else
					nu_contrato="00"&nu_contrato					
				end if
			else
				nu_contrato="0"&nu_contrato					
			end if
		end if	 
		
		
		concatena_contrato = ano_contrato&"/"&nu_contrato
		matricula_contrato = RSC("CO_Matricula")
		situac_contrato = RSC("ST_Contrato")
		
		if situac_contrato="A" then
			situac_contrato_nome="Ativo"		
		else
			situac_contrato_nome="Cancelado"			
		end if		
		
		Set RSA = Server.CreateObject("ADODB.Recordset")	
		SQLA = "SELECT a.NO_Aluno as NOME, m.NU_Unidade as UNI, m.CO_Curso as CUR, m.CO_Etapa as ETA, m.CO_Turma as TUR FROM TB_Alunos a, TB_Matriculas m where a.CO_Matricula = "& matricula_contrato & " and a.CO_Matricula = m.CO_Matricula and m.NU_Ano = "&ano_contrato	
		RSA.Open SQLA, CONa		
			
		if RSA.EOF then
			nome_contrato = "N&atilde;o Informado"
			unidade_contrato = "N.I."
			curso_contrato = "N.I."
			etapa_contrato = "N.I."
			turma_contrato = "N.I."
		else
			nome_contrato = RSA("NOME")		
			co_unidade_contrato = RSA("UNI")
			co_curso_contrato = RSA("CUR")
			etapa_contrato = RSA("ETA")
			turma_contrato = RSA("TUR")	
			
			Set RS0 = Server.CreateObject("ADODB.Recordset")
			SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="&co_unidade_contrato
			RS0.Open SQL0, CON0
			
			unidade_contrato = RS0("NO_Abr")
	
			sql_cu="(Curso='"&curso&"' OR (Curso  is null)) AND"
			Set RS0c = Server.CreateObject("ADODB.Recordset")
			SQL0c = "SELECT * FROM TB_Curso where CO_Curso='"&co_curso_contrato&"'"
			RS0c.Open SQL0c, CON0
			
			curso_contrato = RS0c("NO_Abreviado_Curso")			
		end if	
		
		Set RSB= Server.CreateObject("ADODB.Recordset")
		SQLB = "SELECT * FROM TB_Contrato_Bolsas where CO_Matricula ="&matricula_contrato&" AND NU_Ano_Letivo="&ano_contrato&" AND NU_Contrato = "&nu_contrato
		RSB.Open SQLB, CON5	
		
		if RSB.EOF then	
			bolsista="N&atilde;o"	
		else
			bolsista="Sim"			
		end if	
		

%>
              <tr>
                <td class="<%response.Write(cor)%>"></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(data_contrato)%></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(concatena_contrato)%></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(matricula_contrato)%></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(nome_contrato)%></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(bolsista)%></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(unidade_contrato)%></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(curso_contrato)%></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(etapa_contrato)%></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(turma_contrato)%></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(situac_contrato_nome)%></td>
              </tr>
              <% 		
		intrec=intrec+1
		check=check+1	
	RSC.MOVENEXT
	WEND
end if
%>
              <tr>
                <td colspan="11" align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="50%"><hr></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td><div align="center"> 
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td width="33%"> <div align="center"> 
                        <input name="SUBMIT5" type=button class="botao_cancelar" onClick="MM_goToURL('parent','contratos.asp?pagina=1&v=s');return document.MM_returnValue" value="Voltar">
                    </div></td>
                    <td width="34%"> <div align="center"> </div> <div align="center"> </div></td>
                    <td width="33%"> <div align="center"> 
                        <input name="Submit" type="submit" class="botao_prosseguir" value="Confirmar">
                    </div></td>
                  </tr>
                  <tr>
                    <td width="33%">&nbsp;</td>
                    <td width="34%">&nbsp;</td>
                    <td width="33%">&nbsp;</td>
                  </tr>
                </table>
            <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> </font></div></td>
          </tr>
        </table></td>
    </tr>
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
</form>
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