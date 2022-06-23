<html>
<head>
<title>Escola Bretanha</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/JavaScript">
<!--
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
//-->
</script>
</head>
                    <%
				
usuario = "A"					
Set CON0 = Server.CreateObject("ADODB.Connection") 
CAMINHO0 = "e:\home\bretanha\dados\logins.mdb"
ABRIR0 = "DBQ="& CAMINHO0 & ";Driver={Microsoft Access Driver (*.mdb)}"
CON0.Open ABRIR0
Set RS0 = Server.CreateObject("ADODB.Recordset")
CONEXAO0 = "SELECT * FROM logins WHERE CO_CPF='" & Session("cpf") & "' AND TP_Usuario='" & usuario & "'"
RS0.Open CONEXAO0, CON0
co_mat = RS0("CO_Matricula_Escola")
nome = RS0("NO_Aluno")
%>
<body text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#FFFFFF">
<table width="776" height="100" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="131" valign="top" bgcolor="#6699CC"> 
      <table width=131 border=0 cellpadding=0 cellspacing=0 height="100">
        <tr> 
          <td><a href="../principal_home.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image42','','../images/bot_inicio_dw.gif',1)" target="Conteudo"><img name="Image42" border="0" src="../images/bot_inicio.gif" width="131" height="17"></a></td>
        </tr>
        <tr> 
          <td> <a href="../principal_historia.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image43','','../images/bot_sobre_dw.gif',1)" target="Conteudo"><img name="Image43" border="0" src="../images/bot_sobre.gif" width="131" height="15"></a></td>
        </tr>
        <tr> 
          <td> <a href="../principal_cursos.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image44','','../images/bot_historia_dw.gif',1)" target="Conteudo"><img name="Image44" border="0" src="../images/bot_historia.gif" width="131" height="16"></a></td>
        </tr>
        <tr> 
          <td> <a href="../principal_tour.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image45','','../images/bot_tour_dw.gif',1)" target="Conteudo"><img name="Image45" border="0" src="../images/bot_tour.gif" width="131" height="16"></a></td>
        </tr>
        <tr> 
          <td> <a href="../principal_chegar.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image46','','../images/bot_como_dw.gif',1)" target="Conteudo"><img name="Image46" border="0" src="../images/bot_como.gif" width="131" height="18"></a></td>
        </tr>
        <tr> 
          <td> <a href="../principal_direcao.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image47','','../images/bot_direcao_dw.gif',1)" target="Conteudo"><img name="Image47" border="0" src="../images/bot_direcao.gif" width="131" height="18"></a></td>
        </tr>
        <tr> 
          <td> <a href="../principal_equipe.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image48','','../images/bot_equipe_dw.gif',1)" target="Conteudo"><img name="Image48" border="0" src="../images/bot_equipe.gif" width="131" height="17"></a></td>
        </tr>
        <tr> 
          <td> <a href="../principal_proposta.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image49','','../images/bot_proposta_dw.gif',1)" target="Conteudo"><img name="Image49" border="0" src="../images/bot_proposta.gif" width="131" height="16"></a></td>
        </tr>
        <tr> 
          <td></td>
        </tr>
        <tr> 
          <td> <a href="../principal_calendario.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image50','','../images/bot_cescolar_dw.gif',1)" target="Conteudo"><img name="Image50" border="0" src="../images/bot_cescolar.gif" width="131" height="18"></a></td>
        </tr>
        <tr> 
          <td> <a href="../principal_horarios.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image51','','../images/bot_horarios_dw.gif',1)" target="Conteudo"><img name="Image51" border="0" src="../images/bot_horarios.gif" width="131" height="17"></a></td>
        </tr>
        <tr> 
          <td> <a href="../principal_uniformes.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image52','','../images/bot_unif_dw.gif',1)" target="Conteudo"><img name="Image52" border="0" src="../images/bot_unif.gif" width="131" height="17"></a></td>
        </tr>
        <tr> 
          <td> <a href="../principal_matriculas.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image53','','../images/bot_matricula_dw.gif',1)" target="Conteudo"><img name="Image53" border="0" src="../images/bot_matricula.gif" width="131" height="17"></a></td>
        </tr>
        <tr> 
          <td> <a href="../principal_atividade.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image54','','../images/bot_atividade_dw.gif',1)" target="Conteudo"><img name="Image54" border="0" src="../images/bot_atividade.gif" width="131" height="17"></a></td>
        </tr>
        <tr> 
          <td> <a href="../principal_eventos.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image55','','../images/bot_evento_dw.gif',1)" target="Conteudo"><img name="Image55" border="0" src="../images/bot_evento.gif" width="131" height="17"></a></td>
        </tr>
        <tr> 
          <td> <a href="../principal_contato.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image56','','../images/bot_contato_dw.gif',1)" target="Conteudo"><img name="Image56" border="0" src="../images/bot_contato.gif" width="131" height="16"></a></td>
        </tr>
        <tr> 
          <td> <img src="../images/so_tarja.gif" width=131 height=19 alt=""></td>
        </tr>
         <% 
if   Session("login") = "" then
%>        <tr> 
          <td><a href="../negado.htm" target="Conteudo" onMouseOver="MM_swapImage('Image32','','../images/botsec_boletim_dw.gif',1)" onMouseOut="MM_swapImgRestore()"><img name="Image32" border="0" src="../images/botsec_boletim.gif" width="131" height="18"></a></td>
        </tr>
        <tr> 
          <td><a href="../negado.htm" target="Conteudo" onMouseOver="MM_swapImage('Image31','','../images/botsec_circular_dw.gif',1)" onMouseOut="MM_swapImgRestore()"><img name="Image31" border="0" src="../images/botsec_circular.gif" width="131" height="17"></a></td>
        </tr>
        <tr> 
          <td><a href="../negado.htm" target="Conteudo" onMouseOver="MM_swapImage('Image331','','../images/botsec_ocorre_dw.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="../images/botsec_ocorre.gif" name="Image331" width="131" height="17" border="0" id="Image331"></a></td>
        </tr>
        <tr> 
          <td><a href="../negado.htm" target="Conteudo" onMouseOver="MM_swapImage('Image34','','../images/botsec_financ_dw.gif',1)" onMouseOut="MM_swapImgRestore()"><img name="Image34" border="0" src="../images/botsec_financ.gif" width="131" height="18"></a></td>
        </tr>
        <% Else
%>     <tr> 
          <td><a href="boletim.asp" target="Conteudo" onMouseOver="MM_swapImage('Image32','','../images/botsec_boletim_dw.gif',1)" onMouseOut="MM_swapImgRestore()"><img name="Image32" border="0" src="../images/botsec_boletim.gif" width="131" height="18"></a></td>
        </tr>
        <tr> 
          <td><a href="avisos/default.asp" target="Conteudo" onMouseOver="MM_swapImage('Image31','','../images/botsec_circular_dw.gif',1)" onMouseOut="MM_swapImgRestore()"><img name="Image31" border="0" src="../images/botsec_circular.gif" width="131" height="17"></a></td>
        </tr>
        <%
user = Session("usuario")
if (user ="R") Then
%>        <tr> 
          <td><a href="ocorrencias.asp" target="Conteudo" onMouseOver="MM_swapImage('Image33','','../images/botsec_ocorre_dw.gif',1)" onMouseOut="MM_swapImgRestore()"><img name="Image33" border="0" src="../images/botsec_ocorre.gif" width="131" height="17"></a></td>
        </tr>
        <tr> 
          <td><a href="posicao.asp" target="Conteudo" onMouseOver="MM_swapImage('Image34','','../images/botsec_financ_dw.gif',1)" onMouseOut="MM_swapImgRestore()"><img name="Image34" border="0" src="../images/botsec_financ.gif" width="131" height="18"></a></td>
        </tr>
        <%
else
%>
        <tr> 
          <td height="19">&nbsp; </td>
        </tr>
        <tr> 
          <td height="20">&nbsp; </td>
        </tr>
        <%
End IF
%>
        <% END IF %>

        <tr> 
          <td>&nbsp; </td>
        </tr>
        <tr> 
          <td height="150">&nbsp; </td>
        </tr>
        <tr> 
          <td valign="bottom"><img src="../images/base_fundo_menu.gif" width="131" height="27"></td>
        </tr>
      </table>
















































    </td>
    <td rowspan="2" valign="top"> 
      <table width="635" border="0" cellspacing="8" cellpadding="0">
        <tr> 
          <td width="639" height="29"> 
            <div align="left"><img src="../images/posicao_fin.gif" width="606" height="25"></div></td>
        </tr>
        <tr> 
          <td valign="top" bgcolor="#FFFFFF"> 
            <div align="justify"> 
              <p> 
                <% 'RS_total = RS0.RecordCount 
'RESPONSE.WRITE(RS_total)
				'if (RS_total = 1) then 
				'response.Redirect("boletim2.asp?opt='" & co_mat "'")
'else
While Not RS0.EOF %>
                <font size="2" face="Verdana, Arial, Helvetica, sans-serif"> <img src="../images/seta.jpg" width="48" height="13" align="absbottom"> 
                <a href="posicao2.asp?opt=<%= RS0("CO_Matricula_Escola") %>"><b> 
                <% response.write (RS0("NO_Aluno"))%>
                </b> </a></font> <br>
                <%	 RS0.MoveNext
	Wend 
'END IF
	%></font>
                </p>
            </div></td>
        </tr>
      </table>
    </td>
    <td width="2"></td>
  </tr>
</table>
<table width="761" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#000066" height="11"> 
      <div align="center">
        <a href="../principal_privacidade.asp"><font size="1" color="#FFFFFF" face="Trebuchet MS"><font face="Verdana, Arial, Helvetica, sans-serif"> Pol&iacute;tica de privacidade</font></font></a></div>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#CC0000"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF">Escola 
        Bretanha&copy; - Todos os direitos reservados - 2003 :: Desenvolvido por 
        </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><a href="http://www.simplynet.com.br" target="_blank"> <font color="#FFFFFF">Simply 
        Net</font></a></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"> 
        Informa&ccedil;&atilde;o e Tecnologi</font><font size="1" face="Trebuchet MS" color="#FFFFFF">a</font></div>
    </td>
  </tr>
</table>
</body>

</html>
