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
obr_acc=Session("obr_AAC")
incl_acc=Session("incl_AAC")
Session("obr_AAC")=obr_acc
Session("incl_AAC")=incl_acc

vetor_historico = request.QueryString("dad")
historicos = 	split(vetor_historico,",")

	

Set CON = Server.CreateObject("ADODB.Connection") 
ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
CON.Open ABRIR

Set CON1 = Server.CreateObject("ADODB.Connection") 
ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
CON1.Open ABRIR1

Set CON2 = Server.CreateObject("ADODB.Connection") 
ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
CON2.Open ABRIR2

Set CON7 = Server.CreateObject("ADODB.Connection") 
ABRIR7 = "DBQ="& CAMINHO_h & ";Driver={Microsoft Access Driver (*.mdb)}"
CON7.Open ABRIR7		
		
Set CON0 = Server.CreateObject("ADODB.Connection") 
ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
CON0.Open ABRIR0		

call navegacao (CON,chave,nivel)
navega=Session("caminho")	

ordena = Session("ordena_AEE")

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
			  
        <form action="bd.asp?opt=exc" method="post" name="busca" id="busca" >
          <tr>
      <td valign="top"> 
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>            <tr> 
              
    <td width="766" height="10" colspan="4" valign="top"> 
      <%call mensagens(nivel,421,0,0) %>
    </td>
			  </tr>
          <tr> 
            <td height="10" class="tb_tit"
>Hist&oacute;ricos a serem exclu&iacute;dos</td>
          </tr>
          <tr> 
            <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
              <tr class="tb_subtit">
                <td width="20" height="10">&nbsp;</td>
                <td width="60" align="center">Ano<br>
                  Letivo</td>
                <td width="30" align="center">Seq</td>
                <td width="60" align="center">Matr&iacute;cula</td>
                <td width="280" align="left">&nbsp;Nome</td>
                <td width="40" align="center">Curso</td>
                <td width="110" align="center">Etapa</td>
                <td width="180" align="center">Escola</td>
                <td width="80" align="center">Situa&ccedil;&atilde;o</td>
                <td width="80" align="center">Tipo de Registro</td>
                <td width="60" align="center">Alterado em</td>
              </tr>
              <tr>
                <td colspan="11"><hr width="1000"><input name="exclui_historico" type="hidden" value="<%response.write(vetor_historico)%>"></td>
              </tr>
              <%	
	Set Rs_ordena = Server.CreateObject ( "ADODB.RecordSet" )
	'200 -> VarChar (String), 7 -> Data, 139 -> Numeric
	Rs_ordena.Fields.Append "ano_historico", 139, 10
	Rs_ordena.Fields.Append "seq", 139, 10	
	Rs_ordena.Fields.Append "matric", 139, 10
	Rs_ordena.Fields.Append "nome", 200, 255		
	Rs_ordena.Fields.Append "curso", 200, 3
	Rs_ordena.Fields.Append "segmento", 200, 5
	Rs_ordena.Fields.Append "escola", 200, 50
	Rs_ordena.Fields.Append "situac", 200, 1
	Rs_ordena.Fields.Append "tipo_regist", 200, 1
	Rs_ordena.Fields.Append "data", 7
	Rs_ordena.Open			  
		for de = 0 to ubound(historicos)

			dados_historico = 	split(historicos(de),"$!$")	

			ano_historico = dados_historico(0)
			nu_seq_hist = dados_historico(1)
			cod_aluno  = dados_historico(2)				  
			registros_encontrados = "N"	
			Set RS = Server.CreateObject("ADODB.Recordset")					
			SQL = "SELECT * FROM TB_Historico_Ano where CO_Matricula = "& cod_aluno &" AND DA_Ano = "& ano_historico&" AND NU_Seq = "& nu_seq_hist
			RS.Open SQL, CON7


			if RS.EOF then
				registros_encontrados = "N"	
	%>
            <%else
				registros_encontrados = "S"	
				
				ano_hist = RS("DA_Ano")		 
				seq_hist = RS("NU_Seq")
				matric_hist = RS("CO_Matricula") 
				curso_hist = RS("TP_Curso")
				seg_hist = RS("CO_Seg")
				escola_hist = RS("NO_Escola")
				situac_hist = RS("IN_Aprovado")
				tp_reg_hist = RS("TP_Registro")
				data_hist = RS("DT_Registro")				 
				 
				if data_hist = "" or isnull(data_hist) then
					data_hist = "31/12/9999"
				end if
				 
				Set RSN = Server.CreateObject("ADODB.Recordset")
				SQLN = "SELECT * FROM TB_Alunos where CO_Matricula ="& matric_hist
				RSN.Open SQLN, CON1	
				if RSN.eof then
					nome_hist = "ZZZN&atilde;o cadastrado"
				else
					nome_hist = RSN("NO_Aluno")
				end if		
				
				Rs_ordena.AddNew			
				Rs_ordena.Fields("ano_historico").Value = ano_hist			
				Rs_ordena.Fields("seq").Value = seq_hist
				Rs_ordena.Fields("matric").Value = matric_hist
				Rs_ordena.Fields("nome").Value = nome_hist
				Rs_ordena.Fields("curso").Value = curso_hist
				Rs_ordena.Fields("segmento").Value = seg_hist
				Rs_ordena.Fields("escola").Value = escola_hist
				Rs_ordena.Fields("situac").Value = situac_hist
				Rs_ordena.Fields("tipo_regist").Value = tp_reg_hist
				Rs_ordena.Fields("data").Value = data_hist	
				tot_rec=tot_rec+1	
			end if   	
		NEXT					
							 		
			if ordena="na" then
				Rs_ordena.Sort = "nome ASC"
			elseif ordena="es" then
				Rs_ordena.Sort = "escola ASC"	
			elseif ordena="mt" then
				Rs_ordena.Sort = "matric ASC"	
			else
				Rs_ordena.Sort = "ano_historico ASC"								
			end if
			'Rs_ordena.PageSize = 30
'			 
'			if Request.QueryString("pagina")="" then
'				  intpagina = 1
'				  Rs_ordena.MoveFirst
'			else
'				if cint(Request.QueryString("pagina"))<1 then
'					intpagina = 1
'				else
'					if cint(Request.QueryString("pagina"))>Rs_ordena.PageCount then  
'						intpagina = Rs_ordena.PageCount
'					else
'						intpagina = Request.QueryString("pagina")
'					end if
'				end if   
'			 end if   
'		
'			Rs_ordena.AbsolutePage = intpagina
'			intrec = 0
'			check=2
'			While intrec<Rs_ordena.PageSize and Not Rs_ordena.EoF
			While Not Rs_ordena.EoF
				if check mod 2 =0 then
					cor = "tb_fundo_linha_par" 
				else 
					cor ="tb_fundo_linha_impar"
				end if	
				 
				 
				ano_exibe = Rs_ordena.Fields("ano_historico").Value 			
				seq_exibe = Rs_ordena.Fields("seq").Value 
				matric_exibe = Rs_ordena.Fields("matric").Value 
				nome_exibe = Rs_ordena.Fields("nome").Value 
				curso_exibe = Rs_ordena.Fields("curso").Value 
				seg_bd = Rs_ordena.Fields("segmento").Value 
				escola_exibe = Rs_ordena.Fields("escola").Value 
				situac_exibe = Rs_ordena.Fields("situac").Value 
				tp_reg_exibe = Rs_ordena.Fields("tipo_regist").Value 
				data_exibe = Rs_ordena.Fields("data").Value 		
				
				if isnull(curso_exibe) or isnull(seg_exibe) then
				
				else
					Set RS0 = Server.CreateObject("ADODB.Recordset")
					SQL0 = "SELECT * FROM TB_Segmento where TP_Curso='"&curso_exibe&"' AND CO_Seg='"&seg_bd&"' order by NU_Ordem"
					RS0.Open SQL0, CON7
					seg_exibe = RS0("NO_Abreviado_Curso")	
				end if	
				
				if isnull(situac_exibe) or isnull(situac_exibe) then
				
				else
					Set RS0 = Server.CreateObject("ADODB.Recordset")
					SQL0 = "SELECT * FROM TB_Resultado_Final where TP_Resultado="&situac_exibe
					RS0.Open SQL0, CON7
					situac_exibe = RS0("NO_Resultado")	
				end if								
				
				if tp_reg_exibe = "M" then
					tp_reg_exibe = "Manual"
				else
					tp_reg_exibe = "Autom&aacute;tico"				 	
				end if							 
				 
				data_split= Split(data_exibe,"/")
				dia_s=data_split(0)
				mes_s=data_split(1)
				ano_s=data_split(2)
				
				
				dia=dia*1
				mes=mes*1
				
				if dia_s<10 then
				dia_s="0"&dia_s
				end if
				if mes_s<10 then
				mes_s="0"&mes_s
				end if
			
				da_show=dia_s&"/"&mes_s&"/"&ano_s
			 if da_show = "31/12/9999" then
			 	da_show = ""
			 end if		
		 	if left(nome_exibe,3) = "ZZZ" then
				nome_exibe = replace(nome_exibe,"ZZZ","")
			end if
		 %>
              <tr class="<%=cor%>">
                <td width="20"></td>
                <td width="60" align="center"><%response.Write(ano_exibe)%></td>
                <td width="30" align="center"><%response.Write(seq_exibe)%></td>
                <td width="60" align="center"><%response.Write(matric_exibe)%></td>
                <td width="280" align="left"><%response.Write(nome_exibe)%></td>
                <td width="40" align="center"><%response.Write(curso_exibe)%></td>
                <td width="110" align="center"><%response.Write(seg_exibe)%>
                  <div align="left"></div></td>
                <td width="180" align="center"><%response.Write(escola_exibe)%></td>
                <td width="80" align="center"><%response.Write(situac_exibe)%></td>
                <td width="80" align="center"><%response.Write(tp_reg_exibe)%></td>
                <td width="60" align="center"><%response.Write(da_show)%></td>
              </tr>
              <%check = check+1
				intrec = intrec+1
			Rs_ordena.Movenext
			WEND
%>
            </table></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></div></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td><hr></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td><div align="center"> 
                <table width="1000" border="0" align="center" cellspacing="0">
                  <tr> 
                    <td width="391"> <div align="center"> 
                        <input name="alterar" type="button" class="botao_cancelar" id="alterar" onClick="MM_goToURL('parent','resumo.asp?voltar=S');" value="Voltar">
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