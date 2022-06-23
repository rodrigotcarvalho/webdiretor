<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%
session("nvg")=""
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
chave=request.QueryString("nvg")
opt = request.QueryString("opt")
erro=request.QueryString("res")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
Session("data_consulta")=""
Session("hora_consulta")=""

		
		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2	

 call VerificaAcesso (CON,chave,nivel)
autoriza=Session("autoriza")

 call navegacao (CON,chave,nivel)
navega=Session("caminho")


	%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
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
function submitfuncao()  
{
   var f=document.forms[3]; 
      f.submit(); 
}  function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif"leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<%call cabecalho(nivel)%>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
                  <tr>                    
            
    <td height="10" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> </td>
  </tr>
      <%
	  if autoriza="no" then
	%>
  <tr>                   
    <td height="10"> 
	 <%call mensagens(4,9700,1,0)%>
    </td>
  </tr>	    
    <% else
	
			if opt="err1" then%> 
				  <tr>                   
					<td height="10"> 
					<%call mensagens(4,648,1,0) %>
					</td>
				  </tr>	    
			<%elseif opt="err2" then%>     
				  <tr>                   
					<td height="10"> 
					<%call mensagens(4,649,1,0) %>
					</td>
				  </tr>	
			<%elseif opt="err3" then%>     
				  <tr>                   
					<td height="10"> 
					<%call mensagens(4,650,1,erro) %>
					</td>
				  </tr>	  
			<%elseif opt="err4" then%>     
				  <tr>                   
					<td height="10"> 
					<%call mensagens(4,651,1,erro) %>
					</td>
				  </tr>	                                      			    				  				  
			<% elseif opt="ok" then%>
				  <tr>                   
					<td height="10"> 
					<%call mensagens(4,647,2,0) %>
					</td>
				  </tr>	
			 <%end if
		 end if%>  
      <tr>                   
        <td height="10"> 
        <%call mensagens(4,646,0,0) %>
        </td>
      </tr>	         
  <tr> 
    <td valign="top">

<table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">

          <tr> 
            <td> 
              <%	  if autoriza="no" then			
		else
ano_slct = DatePart("yyyy", now)
mes_slct = DatePart("m", now) 
dia_slct = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

hora = hora*1
min = min*1		
%>
        
      <table width="1000" border="0" cellspacing="0">
        <tr> 
                <td valign="top"><FORM METHOD="POST" ENCTYPE="multipart/form-data" ACTION="upload.asp?opt=f&al=<%=ano_letivo%>" target="_parent">
                  <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo"
>
                    <tr> 
                      <td class="tb_tit">Selecione o arquivo com o resultado</td>
                </tr>
                <tr> 
                  <td> 
                    <table width="1000" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
                      <tr>
                        <td  width="350"><div align="right"><font class="form_dado_texto">Excel: </font></div></td>
                        <td width="650"><INPUT TYPE=FILE SIZE=60 NAME="FILE1" class="borda"></td>
                      </tr>
                      <!-- <tr> 
                  <td width="350"> <div align="right"><font class="form_dado_texto">Arquivo 
                      2: </font></div></td>
                  <td width="650"> <INPUT TYPE=FILE SIZE=60 NAME="FILE2" class="borda"></td>
                </tr>
                <tr> 
                  <td width="350"> <div align="right"><font class="form_dado_texto">Arquivo 
                      3: </font></div></td>
                  <td width="650"> <INPUT TYPE=FILE SIZE=60 NAME="FILE3" class="borda"></td>
                </tr>
                <tr> 
                  <td width="350"> <div align="right"><font class="form_dado_texto">Arquivo 
                      4: </font></div></td>
                  <td width="650"> <INPUT TYPE=FILE SIZE=60 NAME="FILE4" class="borda"></td>
                </tr>
                <tr> 
                  <td width="350"> <div align="right"><font class="form_dado_texto">Arquivo 
                      5: </font></div></td>
                  <td width="650"> <INPUT TYPE=FILE SIZE=60 NAME="FILE5" class="borda"></td>
                </tr> -->
                      <tr>
                        <td colspan="2"><hr width="1000"></td>
                      </tr>
                      <tr>
                        <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="33%"><div align="center"></div></td>
                            <td width="34%"><div align="center"></div></td>
                            <td width="33%"><div align="center">
                              <input name="SUBMIT" type=SUBMIT class="botao_prosseguir" value="Upload!">
                            </div></td>
                          </tr>
                          <tr>
                            <td>&nbsp;</td>
                            <td></td>
                            <td>&nbsp;</td>
                          </tr>
                        </table></td>
                      </tr>
                    </table></td>
                </tr>
              </table></form>
		        </td>
        </tr>
      </table>
      </div> 
      <%end if 
%>
            </td>
          </tr>
      </table>        
      </form></td>
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
response.redirect("../../../../inc/erro.asp")
end if
%>