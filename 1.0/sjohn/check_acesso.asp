<!--#include file="inc/caminhos.asp"-->
<!--#include file="inc/funcoes.asp"-->
<!--#include file="../global/funcoes_diversas.asp"-->
<%
opt=request.QueryString("opt")

	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	

		
if opt="l" then
	dados_cod=request.QueryString("dd")
	dados_decod=Base64Decode(dados_cod)

	dados=SPLIT(dados_decod,"$!$")
	usr_decod=dados(0)
	email_decod=dados(1)
	
	Set RS = Server.CreateObject("ADODB.Recordset")			
	SQL = "select * from TB_Usuario where CO_Usuario = " & usr_decod 
	RS.Open SQL, CON	
	
	if RS.EOF then
		erro="Dados Inválidos"
	else		
		email_bd=RS("TX_EMail_Usuario")
		

		
		if email_decod<>email_bd then
			erro="Dados Inválidos"			
		else	
		
		acesso=1
		ano = DatePart("yyyy", now)
		mes = DatePart("m", now) 
		dia = DatePart("d", now) 
		hora = DatePart("h", now) 
		min = DatePart("n", now) 

		data = dia &"/"& mes &"/"& ano
		horario = hora & ":"& min		
		
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2= "UPDATE TB_Usuario SET NU_Acesso= "& acesso & ", HO_ult_Acesso = '"& horario & "', DA_Ult_Acesso = '"& data & "', ST_Usuario='L' WHERE CO_Usuario = "&usr_decod
		RS2.Open SQL2, CON		
		
		'response.redirect ("inicio.asp?opt=sa")
			response.redirect ("default.asp")
		end if
	end if	
else

	login=request.QueryString("lg")
	
	Set RS = Server.CreateObject("ADODB.Recordset")			
	SQL = "select * from TB_Usuario where CO_Usuario = " & login 
	RS.Open SQL, CON		

	email_bd=RS("TX_EMail_Usuario")

	nivel=999
	%>
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <title>Web Família</title>
    <link href="estilo.css" rel="stylesheet" type="text/css" />
    <script type="text/JavaScript">
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
    function MM_swapImgRestore() { //v3.0
      var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
    }
    
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
    
    function MM_swapImage() { //v3.0
      var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
       if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
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
    <script language="JavaScript" type="text/JavaScript">
    <!--
    function MM_reloadPage(init) {  //reloads the window if Nav4 resized
      if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
        document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
      else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
    }
    MM_reloadPage(true);
    
    function submit()  
    {
       var f=document.forms[0]; 
          f.submit(); 
    }
    
    function MM_callJS(jsStr) { //v2.0
      return eval(jsStr)
    }
    //-->
    </script>
    
    </head>
    
    <body onLoad="MM_preloadImages(<%response.Write(swapload)%>)">
    <table width="1000" height="500" border="0" align="center" cellpadding="0" cellspacing="0" class="tabela_principal">
      <tr>
        <td height="500"><table width="200" height="600" border="0" cellpadding="0" cellspacing="0">
            <!--DWLayoutTable-->
            <tr valign="bottom"> 
              <td height="90"> 
                <%call cabecalho(nivel)%>
              </td>
            </tr>
            <tr class="tabela_menu">
              <td height="5" valign="top"><p><img src="img/linha-pontilhada_grande.gif" alt="" width="828" height="5" /></p></td> 
            </tr>
          <tr class="tabela_menu">
            <td height="405" valign="top"> <table width="70%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="center"><div class="textbox"><p class="tb_fundo_linha_ok">Foi encaminhada uma mensagem para a caixa postal <%response.Write(email_bd)%>.<br>Clique no link presente na mensagem para liberar o acesso ao Web Família.<br><br> 
* Caso não receba o e-mail, verifique se o mesmo não se encontra em sua caixa de spam.<br>
** Se não estiver na caixa de spam, entre em contato com a secretaria do colégio.</p></div></td>
  </tr>
</table>
</td> 
            </tr>
          </table></td>
      </tr>
      <tr>
        <td width="1000" height="41"><img src="img/rodape.jpg" width="1000" height="41" /></td>
      </tr>
    </table>
    </body>
    </html>
<%end if%>        