<%
Function Banner(nivel,this_file,sistema_local,nome,permissao,ano_letivo)
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

co_usr = session("co_user")

'select case mes
' case 1 
' mes = "janeiro"
' case 2 
' mes = "fevereiro"
' case 3 
' mes = "março"
' case 4
' mes = "abril"
' case 5
' mes = "maio"
' case 6 
' mes = "junho"
' case 7
' mes = "julho"
' case 8 
' mes = "agosto"
' case 9 
' mes = "setembro"
' case 10 
' mes = "outubro"
' case 11 
' mes = "novembro"
' case 12 
' mes = "dezembro"
'end select

data = dia &" / "& mes &" / "& ano
horario = hora & ":"& min
horaagora = hour(now)
data= FormatDateTime(data,1) 

if horaagora < 24 then cumprimento = "Boa noite "
if horaagora < 18 then cumprimento = "Boa tarde "
if horaagora < 12 then cumprimento = "Bom dia "

	
		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON_p = Server.CreateObject("ADODB.Connection") 
		ABRIR_p = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_p.Open ABRIR_p

		
select case nivel
case 0
parent_folder="" 
case 1
parent_folder="../"
case 2
parent_folder="../../"
case 3
parent_folder="../../../"
case 4
parent_folder="../../../../"
end select

%>
<table width="1000" height="135" border="0" align="center" cellpadding="0" cellspacing="0" id="Table_01">
	<tr>
		<td>
			<img src="<%response.write(parent_folder)%>img/banner_webdiretor_01.gif" width="250" height="68" alt=""></td>
		<td>
			<img src="<%response.write(parent_folder)%>img/banner_webdiretor_02.gif" width="250" height="68" alt=""></td>
		<td>
			<img src="<%response.write(parent_folder)%>img/banner_webdiretor_03.gif" width="248" height="68" border="0" alt="" usemap="#banner_webdiretor_03_Map"></td>
		<td>
			<img src="<%response.write(parent_folder)%>img/banner_webdiretor_04.gif" width="252" height="68" border="0" alt="" usemap="#banner_webdiretor_04_Map"></td>
	</tr>
	<tr>
		<td>
			<img src="<%response.write(parent_folder)%>img/banner_webdiretor_05.gif" width="250" height="67" alt=""></td>
		<td>
			<img src="<%response.write(parent_folder)%>img/banner_webdiretor_06.gif" width="250" height="67" alt=""></td>
		<td>
			<img src="<%response.write(parent_folder)%>img/banner_webdiretor_07.gif" width="248" height="67" alt=""></td>
		<td>
			<img src="<%response.write(parent_folder)%>img/banner_webdiretor_08.gif" width="252" height="67" alt=""></td>
	</tr>
  <tr class="saudacao_banner"> 
    <td colspan="3">Ol&aacute;<span class="nome_banner"> 
      <%response.Write (nome)%>
      </span>, &uacute;ltimo acesso dia 
      <% Response.Write(session("dia_t")) %>
      &agrave;s 
      <% Response.Write(session("hora_t")) %>
      </span></td>
    <td><div align="right"><span class="data_banner"> 
        <% response.write (data)%>
        </span></div></td>
  </tr>
  <tr> 
    <td colspan="4"><table width="100%" height="14" border="0" cellspacing="0">
          <tr class="menu_banner"> 
   <form name="ano_letivo" action="<%response.write(parent_folder)%>inc/redireciona.asp?opt=al" method="post">		  
            <td width="15%" height="10"> <div align="right">Ano 
                Letivo</div></td>
            <td width="15%" height="10"> 
			<select name="ano_letivo" class="select_style" onChange="MM_callJS('submitano()')">
                <%
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Ano_Letivo order by NU_Ano_Letivo"
		RS0.Open SQL0, CON
		while not RS0.EOF 
		ano_bd=RS0("NU_Ano_Letivo")
		ano_info=nivel&"-"&this_file&"-"&ano_bd
		ano_letivo=ano_letivo*1
		ano_bd=ano_bd*1		
				if ano_letivo=ano_bd then%>
                <option value="<%=ano_info%>" selected><%=ano_bd%></option>
                <%else%>
                <option value="<%=ano_info%>"><%=ano_bd%></option>
                <%end if
		RS0.MOVENEXT
		WEND 		
				%>
              </select> </td>
   </form>			  
	<form name="ano_letivo" action="<%response.write(parent_folder)%>inc/redireciona.asp?opt=sa" method="post">		  
            <td width="15%" height="10"> <div align="right">Sistemas 
                Autorizados </div></td>
            <td width="15%" height="10">  
			<select name="sistema" id="sistema" class="select_style" onChange="MM_callJS('submitsistema()')">
				<%
				nivel=nivel*1
				if nivel=0 then%>
                <option value="0-WR-inicio.asp" selected></option>			
                <%end if



					  
  		  		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Acessos where Permissao = "& permissao&" order by NU_Pos"
		RS.Open SQL, CON
	
			
	

	while not RS.EOF
		
co_sistema=RS("CO_Sistema")

if co_sistema="WR" then
	RS.MOVENEXT
elseif co_sistema="WN" then
		Set RS_p = Server.CreateObject("ADODB.Recordset")
		SQL_p = "SELECT * FROM TB_Professor where CO_Usuario = "&co_usr
		RS_p.Open SQL_p, CON_p
	if RS_p.EOF then
		RS.MOVENEXT
	else
			Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Sistema where CO_Sistema = '"&co_sistema&"' order by NU_Pos"
		RS1.Open SQL1, CON
		
		sistema=RS1("TX_Descricao")
		link=RS1("CO_Pasta")
		info=nivel&"-"&co_sistema&"-"&link
		if co_sistema= sistema_local then
		%>
                <option value="<%=info%>" selected><%=server.HTMLEncode(sistema)%></option>
		<%else%>				
                <option value="<%=info%>"><%=server.HTMLEncode(sistema)%></option>				
        <%end if				
	RS.MOVENEXT
	end if
else		
		
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Sistema where CO_Sistema = '"&co_sistema&"' order by NU_Pos"
		RS1.Open SQL1, CON
		
		sistema=RS1("TX_Descricao")
		link=RS1("CO_Pasta")
		info=nivel&"-"&co_sistema&"-"&link
		if co_sistema= sistema_local then
%>
                <option value="<%=info%>" selected><%=server.HTMLEncode(sistema)%></option>
<%else%>				
                <option value="<%=info%>"><%=server.HTMLEncode(sistema)%></option>				
                <%	
end if				
	RS.MOVENEXT
end if
	WEND%>
              </select> </td>
   </form>			  
   <form name="ano_letivo" action="<%response.write(parent_folder)%>inc/redireciona.asp?opt=ar" method="post">			  
            <td width="15%" height="10"> <div align="right"></div></td>
            <td width="15%" height="14">&nbsp; </td>

   </form>			  
          </tr>
                </table></td>
  </tr>
</table>
<map name="banner_webdiretor_03_Map">
<area shape="rect" alt="Fale Conosco" coords="193,11,248,24" href="<%response.write(parent_folder)%>faleconosco.asp">
<area shape="rect" alt="Home" coords="97,11,176,25" href="<%response.write(parent_folder)%>inicio.asp">
</map>
<map name="banner_webdiretor_04_Map">
<area shape="rect" alt="Sair" coords="215,11,242,24" href="<%response.write(parent_folder)%>sair.asp">
<area shape="rect" alt="Novo Login" coords="136,11,196,24" href="<%response.write(parent_folder)%>default.asp">
<area shape="rect" alt="Alterar Senha" coords="39,11,125,23" href="<%response.write(parent_folder)%>seguranca.asp">
<area shape="rect" alt="Fale Conosco" coords="0,11,20,24" href="<%response.write(parent_folder)%>faleconosco.asp">
</map>
<%end function%>