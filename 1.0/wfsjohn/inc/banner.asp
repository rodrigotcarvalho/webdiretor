<%
Function Banner(nivel,this_file,sistema_local,nome,permissao,ano_letivo)
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

co_usr = session("co_user")

select case mes
 case 1 
 mes = "janeiro"
 case 2 
 mes = "fevereiro"
 case 3 
 mes = "mar&ccedil;o"
 case 4
 mes = "abril"
 case 5
 mes = "maio"
 case 6 
 mes = "junho"
 case 7
 mes = "julho"
 case 8 
 mes = "agosto"
 case 9 
 mes = "setembro"
 case 10 
 mes = "outubro"
 case 11 
 mes = "novembro"
 case 12 
 mes = "dezembro"
end select

data = dia &" / "& mes &" / "& ano
horario = hora & ":"& min
horaagora = hour(now)
data= FormatDateTime(Date(),1) 

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
	link="s"	
case 1
	parent_folder="../"
	link="s"	
case 2
	parent_folder="../../"
	link="s"	
case 3
	parent_folder="../../../"
	link="s"	
case 4
	parent_folder="../../../../"
	link="s"	
case 999
	parent_folder=""
	link="n"
end select
if session("tp")="R" then
inicio="inicio.asp?opt=sa"
else
inicio="inicio.asp?opt=ad"
end if
%>
<table id="Table_01" width="1000" height="135" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td>
			<img src="<%response.write(parent_folder)%>img/banner_webfamilia_01.gif" width="249" height="68" alt=""></td>
		<td>
			<img src="<%response.write(parent_folder)%>img/banner_webfamilia_02.gif" width="251" height="68" alt=""></td>
		<td>
			<img src="<%response.write(parent_folder)%>img/banner_webfamilia_03.gif" width="248" height="68" alt=""></td>
		<td>
			<img src="<%response.write(parent_folder)%>img/banner_webfamilia_04.gif" width="252" height="68" border="0" alt="" usemap="#banner_webfamilia_04_Map"></td>
	</tr>
	<tr>
		<td>
			<img src="<%response.write(parent_folder)%>img/banner_webfamilia_05.gif" width="249" height="67" alt=""></td>
		<td>
			<img src="<%response.write(parent_folder)%>img/banner_webfamilia_06.gif" width="251" height="67" alt=""></td>
		<td>
			<img src="<%response.write(parent_folder)%>img/banner_webfamilia_07.gif" width="248" height="67" alt=""></td>
		<td>
			<img src="<%response.write(parent_folder)%>img/banner_webfamilia_08.gif" width="252" height="67" alt=""></td>
	</tr>
</table>
<%if link="s" then%>
    <map name="banner_webfamilia_04_Map">
    <area shape="rect" alt="" coords="216,10,237,24" href="<%response.write(parent_folder&"default.asp")%>" onclick = "if (! confirm('Você deseja encerrar sua conexão com o Web Família?')) { return false; }">
    <area shape="rect" alt="" coords="121,10,201,24" href="<%response.write(parent_folder&"faleconosco.asp")%>">
    <area shape="rect" alt="" coords="28,10,111,25" href="<%response.write(parent_folder&inicio)%>">
    </map>   
<%else%> 
    <map name="banner_webfamilia_04_Map">
    <area shape="rect" alt="" coords="216,10,237,24" href="<%response.write(parent_folder&"default.asp")%>">
    </map>   
<%end if%>  
<%end function%>
