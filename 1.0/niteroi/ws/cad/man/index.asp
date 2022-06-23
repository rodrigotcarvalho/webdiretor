<%'On Error Resume Next%>
<!--#include file="../../../inc/caminhos.asp"-->
<!--#include file="../../../inc/funcoes.asp"-->
<%
' váriáveis de sessão são capturadas em inc/funcoes.asp
nivel=3
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
chave=request.QueryString("nvg")
ano_info=nivel&"-"&chave&"-"&ano_letivo


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

 call navegacao (CON,chave,nivel)
navega=Session("caminho")

%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="../../../js/global.js"></script>
<script language="JavaScript">
 window.history.forward(1);
</script>
<script language="JavaScript" type="text/JavaScript">
<!--

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
      img.src = (img.MM_dn && args[i+2]) ? args[i+2] : ((args[i+1])?args[i+1] : img.MM_up);
      nbArr[nbArr.length] = img;
    }
  } else if (event == "out" ) {
    for (i=0; i < document.MM_nbOver.length; i++) { img = document.MM_nbOver[i]; img.src = (img.MM_dn) ? img.MM_dn : img.MM_up; }
  } else if (event == "down") {
    nbArr = document[grpName];
    if (nbArr) for (i=0; i < nbArr.length; i++) { img=nbArr[i]; img.src = img.MM_up; img.MM_dn = 0; }
    document[grpName] = nbArr = new Array();
    for (i=2; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
      if (!img.MM_up) img.MM_up = img.src;
      img.src = img.MM_dn = (args[i+1])? args[i+1] : img.MM_up;
      nbArr[nbArr.length] = img;
  } }
}

function MM_preloadImages() { //v3.0
 var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
   var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
   if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}//-->
</script>
<link href="../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="../../../img/fundo.gif" onLoad="MM_preloadImages('../../../img/baner_export_r2_c2_f2.gif','../../../img/baner_export_r2_c2_f4.gif','../../../img/baner_export_r2_c2_f3.gif','../../../img/baner_export_r4_c2_f2.gif','../../../img/baner_export_r4_c2_f4.gif','../../../img/baner_export_r4_c2_f3.gif','../../../img/baner_export_r2_c5_f2.gif','../../../img/baner_export_r2_c5_f4.gif','../../../img/baner_export_r2_c5_f3.gif','../../../img/baner_export_r3_c4_f2.gif','../../../img/baner_export_r3_c4_f4.gif','../../../img/baner_export_r3_c4_f3.gif')">

<%

call cabecalho(nivel)

%>
        
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
            <td valign="top"> 
              <div align="center">
                
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
				  
<%
'		Set RS2 = Server.CreateObject("ADODB.Recordset")
'		SQL2 = "SELECT * FROM TB_Modulo where CO_Sistema='"&sistema_local&"' order by NU_Pos"
'		RS2.Open SQL2, CON
		
'		While not RS2.EOF
'modulo=RS2("CO_Modulo")
'modulo_nome=RS2("TX_Descricao")
'link_modulo=RS2("CO_Pasta")
'modulo_nome= " > <a href='"&link_modulo&"' class='caminho' target='_self'>"&modulo_nome&"</a>"
%>					  
                  <tr> 
                    
            <td height="10" class="tb_caminho">
              <%
	  response.Write(navega&modulo_nome)

%>
              </td>
                  </tr>
          <tr> 
                    
            <td height="10"> 
              <%call mensagens(nivel,0,0,0) %>
            </td>
                  </tr>				  
          <tr> 
                    <td>				  
<table width="100%" border="0" cellspacing="0">

<%		
nvgarray=split(chave,"-")
sistema=nvgarray(0)
modulo=nvgarray(1)
setor=nvgarray(2)
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Modulo where CO_Sistema='"&sistema&"' And CO_Modulo='"&modulo&"'"
		RS2.Open SQL2, CON

modulo=RS2("CO_Modulo")
modulo_nome=RS2("TX_Descricao")
link_modulo=RS2("CO_Pasta")
nvg=sistema_local&"-"&modulo
modulo_nome= "<strong><a href='../index.asp?nvg="&nvg&"' class='modulo' target='_self'>"&modulo_nome&"</a></strong>"
%>					  
                  <tr> 
                    
            <td height="10" class="tb_modulo"> 
              <%
	  response.Write(modulo_nome)

%>
</td>
                  </tr>
<%
Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Setor where CO_Setor='"&setor&"' AND CO_Modulo='"&modulo&"' AND CO_Sistema='"&sistema&"' order by NU_Pos"
		RS3.Open SQL3, CON

linha=1
registro=1

setor_nome=RS3("TX_Descricao")
link_setor=RS3("CO_Pasta")

%>
                  <td width="<%=larg_col%>%" valign="top"> 
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td height="14" class="tb_setor"> 
                          <%response.Write(setor_nome)%>
                        </td>
				   </tr>
				   <tr> 
                        <td class="linkum"> 
                          <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Funcao where CO_Setor='"&setor&"' AND CO_Modulo='"&modulo&"' AND CO_Sistema='"&sistema&"' order by NU_Pos"
		RS4.Open SQL4, CON

		While not RS4.EOF
		nu_pos=RS4("NU_Pos")
		if NU_Pos=0 then
		RS4.Movenext
		else
		funcao=RS4("CO_Funcao")					
		funcao_nome=RS4("TX_Descricao")
		link_funcao=RS4("CO_Pasta")
		funcao_nome= "• <a href='"&link_funcao&"/index.asp?nvg="&sistema_local&"-"&modulo&"-"&setor&"-"&funcao&"' class='linkum' target='_self'>"&funcao_nome&"</a><br>"
		%>
                          <%response.Write(funcao_nome)%>
                          <%		
		RS4.Movenext
		end if
		Wend
		%> </td>
				   </tr>
                      </table>
                  </td>				  				  

                  </tr>
                  			  

</table>				  
</td>
          </tr>				  
		  
                </table>
              </div></td>
          </tr>
          <tr>
            <td height="40" valign="top" bgcolor="#FFFFFF"><img src="../../../img/rodape.jpg" width="1000" height="40"></td>
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
pasta=arPath(seleciona)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../inc/erro.asp")
end if
%>