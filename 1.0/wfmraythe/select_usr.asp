<!--#include file="inc/caminhos.asp"-->
<%
opt = request.QueryString("opt")

	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1		

if opt="on" then
	login =request.form("login")
	if login="" or isnull(login) then
		msg="O campo Usuário é obrigatório!"
		tipo="e"	
	else	
	
		Set RS = Server.CreateObject("ADODB.Recordset")			
		SQL = "select * from TB_Usuario where CO_Usuario = " & login 
		RS.Open SQL, CON

		if RS.eof and RS.bof then 
			msg="O Usuário "&login&" não existe!"
			tipo="e"	
		else
			st_usuario = RS("ST_Usuario")
			tp_usuario = RS("TP_Usuario")
			
			if st_usuario="B" then
				msg="O Usuário "&login&" está bloqueado!"
				tipo="e"	
			else

				co_user= RS("CO_Usuario")
				acesso = RS("NU_Acesso")
				data_de = RS("DA_Ult_Acesso")
				hora_de = RS("HO_ult_Acesso")
				session("nome") = RS("NO_Usuario")
				session("login") = co_user
				session("tp") = RS("TP_Usuario")			
				session("acesso") = acesso
				session("co_user") = co_user
				session("permissao") = permissao
				session("sistema_local")="raiz"
				session("escola")=escola
				session("trava")="n"		
	
				if data_de="" or isnull(data_de) then
					ano_de = DatePart("yyyy", now)
					mes_de = DatePart("m", now) 
					dia_de = DatePart("d", now) 
				else			
					dados_dtd= split(data_de, "/" )
					dia_de= dados_dtd(0)
					mes_de= dados_dtd(1)
					ano_de= dados_dtd(2)
				end if
				
				if hora_de="" or isnull(hora_de) then
					h_de = DatePart("h", now) 
					min_de = DatePart("n", now) 				
				else	
					dados_hrd= split(hora_de, ":" )
					h_de= dados_hrd(0)
					min_de= dados_hrd(1)
				end if
				'dia_de=dia_de*1
				'if dia_de<10 then
				'	dia_de="0"&dia_de
				'end if
				'mes_de=mes_de*1					
				'if mes_de<10 then
				'	mes_de="0"&mes_de
				'end if
				h_de=h_de*1							
				if h_de<10 then
					h_de="0"&h_de
				end if
				hora_de=h_de&":"&min_de
				
				data_inicio=dia_de&"/"&mes_de&"/"&ano_de
				data_inicio=FormatDateTime(data_inicio,2)
				session("hora_t") = hora_de			
	
				Set RSano = Server.CreateObject("ADODB.Recordset")
				SQLano = "SELECT * FROM TB_Ano_Letivo where ST_Ano_Letivo='L' order by NU_Ano_Letivo"
				RSano.Open SQLano, CON
		
				ano_letivo=RSano("NU_Ano_Letivo")
				session("ano_letivo") = ano_letivo
				min_ano_letivo = ano_letivo				
				session("ano_vigente") = ano_letivo						
				
				max_ano_letivo = 0
				if session("tp")="R" then
					Set RS2 = Server.CreateObject("ADODB.Recordset")						
					SQL2 = "select * from TB_RespxAluno where CO_Usuario = " & co_user &" ORDER BY CO_Aluno"
					RS2.Open SQL2, CON						
						
					while not RS2.EOF
						co_aluno=RS2("CO_Aluno")

						Set RS3 = Server.CreateObject("ADODB.Recordset")					
						SQL3 = "select MIN(NU_Ano) as min_ano_letivo, MAX(NU_Ano) as max_ano_letivo from TB_Matriculas where CO_Matricula = " & co_aluno 
					
						RS3.Open SQL3, CON1
						
						temp_min_ano_letivo=RS3("min_ano_letivo")
						temp_min_ano_letivo = temp_min_ano_letivo*1
						min_ano_letivo = min_ano_letivo*1
						if temp_min_ano_letivo<min_ano_letivo  then
							min_ano_letivo = temp_min_ano_letivo
						end if
					
						temp_max_ano_letivo=RS3("max_ano_letivo")
						temp_max_ano_letivo = temp_max_ano_letivo*1
						max_ano_letivo = max_ano_letivo*1
						if temp_max_ano_letivo> max_ano_letivo then
							max_ano_letivo = temp_max_ano_letivo
						end if		
						
'response.write(temp_max_ano_letivo&";"&temp_max_ano_letivo&";"&anos_anteriores_exibidos&";"&min_ano_letivo&";"&max_ano_letivo&"<BR>")									
					RS2.MOVENEXT
					Wend						
					
					ano_letivo=ano_letivo*1	
					menor_ano_exibido = menor_ano_exibido*1				
					menor_ano_exibido = ano_letivo-anos_anteriores_exibidos
'response.write(tp_usuario&"-"&menor_ano_exibido&";"&ano_letivo&";"&anos_anteriores_exibidos&";"&min_ano_letivo&";"&max_ano_letivo&"<BR>")			
					if min_ano_letivo< menor_ano_exibido then
						min_ano_letivo = menor_ano_exibido
					end if	
					max_ano_letivo=max_ano_letivo*1
					ano_letivo=ano_letivo*1
					if max_ano_letivo< ano_letivo then
						session("ano_letivo") = max_ano_letivo
					end if					
'response.write(tp_usuario&"-"&menor_ano_exibido&";"&ano_letivo&";"&anos_anteriores_exibidos&";"&min_ano_letivo&";"&max_ano_letivo&"<BR>")			
'response.end()										
					session("min_ano_letivo") = min_ano_letivo
					session("max_ano_letivo") = max_ano_letivo	
					response.redirect ("inicio.asp?opt=sa")
				else
					Set RS3 = Server.CreateObject("ADODB.Recordset")					
					SQL3 = "select MIN(NU_Ano) as min_ano_letivo, MAX(NU_Ano) as max_ano_letivo from TB_Matriculas where CO_Matricula = " & co_user 
					RS3.Open SQL3, CON1	
					
					min_ano_letivo=RS3("min_ano_letivo")					
					max_ano_letivo=RS3("max_ano_letivo")									
					
					ano_letivo=ano_letivo*1	
					menor_ano_exibido = menor_ano_exibido*1				
					menor_ano_exibido = ano_letivo-anos_anteriores_exibidos
		
					if min_ano_letivo< menor_ano_exibido then
						min_ano_letivo = menor_ano_exibido
					end if	
					max_ano_letivo=max_ano_letivo*1
					ano_letivo=ano_letivo*1
					if max_ano_letivo< ano_letivo then
						session("ano_letivo") = max_ano_letivo
					end if		
									
					session("min_ano_letivo") = min_ano_letivo
					session("max_ano_letivo") = max_ano_letivo			
					response.redirect ("inicio.asp?opt=ad")
				end if
				
										
			end if
		end if
	end if
else
	
end if


if tipo="e" then
	cor = "#FF0000"
elseif tipo="o" then
	cor = "#003399"
end if

%>

<html>
<head>
<title>Web Diretor</title>
<link href="estilo.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
@import url("tabelass.css");
body {
	background-image: url(img/grade-fundo.gif);
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<script language="JavaScript"> 
function FocusNoForm() 
{ 
//formlogin.nome.value="testes"; 
login.login.focus(); 
} 
</script> 
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function createXMLHTTP()
	{
		try
		{
		   ajax = new ActiveXObject("Microsoft.XMLHTTP");
		}
		catch(e)
		{
		   try
		   {
					   ajax = new ActiveXObject("Msxml2.XMLHTTP");
					   alert(ajax);
		   }
		   catch(ex)
		   {
					   try
					   {
								   ajax = new XMLHttpRequest();
					   }
					   catch(exc)
					   {
									alert("Esse browser não tem recursos para uso do Ajax");
									ajax = null;
					   }
		 	}
			return ajax;
		}
	
		var arrSignatures = ["MSXML2.XMLHTTP.5.0", "MSXML2.XMLHTTP.4.0",
		"MSXML2.XMLHTTP.3.0", "MSXML2.XMLHTTP",
		"Microsoft.XMLHTTP"];

		for (var i=0; i < arrSignatures.length; i++) {
														  try {
																 var oRequest = new ActiveXObject(arrSignatures[i]);
																 return oRequest;
															  } catch (oError) {
																			   }
													  }
		
		  throw new Error("MSXML is not installed on your system.");
	}                                
						
						
	 function GuardaLogin(login)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../global/guarda_valores_digitados.asp?opt=valor1", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor1=" + login);
		}
	
	
	 function GuardaSenha(senha)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../global/guarda_valores_digitados.asp?opt=valor2", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor2=" + senha);
		}

 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
</script>
</head>

<body bgcolor="#FFFFFF" topmargin="100" marginheight="100" onLoad="FocusNoForm()">
<form action="select_usr.asp?opt=on" method="post" name="login" id="login" autocomplete="OFF">
          
  <table width="1000" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="10">&nbsp;</td>
  </tr>
  <tr>
    <td><table width="801" height="535" border="0" align="center" cellpadding="0" cellspacing="0" background="img/select_usr.png">
    <tr> 
              <td width="36" height="235">&nbsp;</td>
              <td height="235" colspan="2">&nbsp;</td>
    </tr>
            <tr> 
              <td height="298">&nbsp;</td>
              <td width="373" height="298" valign="top"><span class="texto_link style1">
                <input name="log" type="hidden" id="log" value="on">
              </span></td>
              <td width="392" height="298" valign="top"><table width="339" height="244" border="0" align="left" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="339" height="40"><table width="334" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="25"><table width="81%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td width="41%" align="right"><img src="img/usuario.png" width="47" height="15"></td>
                          <td width="59%" align="right"><input name="login" type="text" class="textbox" id="login" size="25" onKeyUp="GuardaLogin(this.value)"></td>
                        </tr>
                      </table></td>
                      </tr>
                    <tr>
                      <td height="25"><div align="left">
                        <table width="81%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="39%" align="right">&nbsp;</td>
                            <td width="61%" align="right">&nbsp;</td>
                          </tr>
                        </table>
                      </div></td>
                      </tr>
                    <tr>
                      <td height="15"><table width="81%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td width="49%" height="15">&nbsp;</td>
                          <td width="51%" height="15" align="right" valign="top">&nbsp;</td>
                        </tr>
                      </table></td>
                      </tr>
                    <tr>
                      <td height="5"></td>
                      </tr>
                    <tr>
                      <td height="25"><div align="left">
                        <table width="81%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="59%" align="right">&nbsp;</td>
                            <td width="41%" align="right">&nbsp;</td>
                          </tr>
                        </table>
                      </div></td>
                      </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="82">&nbsp;</td>
                </tr>
                <tr>
                  <td height="18" valign="top"><div align="center">
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="54%" height="15">&nbsp;</td>
                        <td width="46%" height="15">&nbsp;</td>
                        </tr>
                    </table>
                  </div></td>
                </tr>
                <tr>
                  <td height="29"><div align="center">
                    <input name="escola" type="hidden" id="escola" value="5">
                    <input name="Enviar" type="image" src="img/botao_autenticar.gif" alt="autenticar" width="130" height="30" border="0">
                  </div></td>
                </tr>
                <tr>
                  <td height="18" valign="top"><table width="337" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="337" height="13"><div align="center"><a href="envia_senha.asp"><img src="img/botao_transparente.gif" alt="esqueci a senha" width="100" height="15" border="0" align="middle"></a></div></td>
                    </tr>
                  </table></td>
                </tr>
              </table></td>
            </tr>		
  </table></td>
  </tr>
</table>

<% if tipo="e" or tipo="o" then%>    
  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td><table width="550" height="40" border="1" align="center" cellpadding="0" cellspacing="1" bordercolor="<%=cor%>">
          <tr bgcolor="<%=cor%>"> 
            <td><div align="center"> <font color="#FFFFFF"><strong><font size="1" face="Arial, Helvetica, sans-serif"> 
                <%response.Write(msg)%>
                </font></strong></font></div></td>
          </tr>
      </table></td>
    </tr>
  </table>
<%end if%>
</form>
</body>
</html>
