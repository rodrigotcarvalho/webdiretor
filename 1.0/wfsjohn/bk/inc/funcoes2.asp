<!--#include file="../inc/connect_pr.asp"-->


<%
Function aniversario(y,m,d)
ano = y
mes = m
dia = d

data= dia&"-"&mes&"-"&ano
intervalo = DateDiff("d", data , now )

intervalo = int(intervalo/365.25)

response.write(intervalo&" anos")

End Function






Function GeraNomes(materia,unidade,curso,etapa,Conexao)

Sqlmt= "SELECT * FROM TB_Materia where CO_Materia = '"& materia&"'"
Set rsmt= Conexao.Execute ( Sqlmt ) 
IF rsmt.eof THEN
no_materia= ""
ELSE
no_materia= rsmt("NO_Materia")
END IF


Sqlun= "SELECT * FROM TB_Unidade where NU_Unidade = "& unidade
Set rsun= Conexao.Execute ( Sqlun ) 
IF rsun.eof THEN
no_unidade= ""
ELSE
no_unidade= rsun("NO_Unidade")
END IF



Sqlgr= "SELECT * FROM TB_Curso where CO_Curso = '"& curso &"'"
Set rsgr= Conexao.Execute ( Sqlgr ) 
IF rsgr.eof THEN
no_curso= ""
ELSE
no_curso= rsgr("NO_Curso")
prep_curso= rsgr("CO_Conc")
END IF


Sqlsr= "SELECT * FROM TB_Etapa where CO_Curso = '"& curso &"' and CO_Etapa = '"& etapa &"'"
Set rssr= Conexao.Execute ( Sqlsr ) 
IF rssr.eof THEN
no_etapa= ""
ELSE
no_etapa= rssr("NO_Etapa")
END IF

session("no_materia") = no_materia
session("no_unidade") = no_unidade
session("prep_curso") = prep_curso
session("no_curso") = no_curso
session("no_etapa") = no_etapa

end Function

Function GeraNomesMapao(unidades,grau,serie,Conexao)

Sqlun= "SELECT * FROM TB_Unidade where NU_Unidade = "& unidades
Set rsun= Conexao.Execute ( Sqlun ) 
no_unidades= rsun("NO_Unidade")

Sqlgr= "SELECT * FROM TB_Curso where CO_Curso = '"& grau &"'"
Set rsgr= Conexao.Execute ( Sqlgr ) 
no_grau= rsgr("NO_Curso")

Sqlsr= "SELECT * FROM TB_Etapa where CO_Curso = '"& grau &"' and CO_Etapa = '"& serie &"'"
Set rssr= Conexao.Execute ( Sqlsr ) 
no_serie= rssr("NO_Etapa")

session("no_materia") = no_materia
session("no_unidades") = no_unidades
session("no_grau") = no_grau
session("no_serie") = no_serie

end Function

'///////////////////////////////////////////////    Último  //////////////////////////////////////////////////////////////

FUNCTION ultimo(tb)

session("codigo_u") = 0
session("codigo_u2") = 0
select case tb

case 0


		Set CONu = Server.CreateObject("ADODB.Connection") 
		ABRIRu = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONu.Open ABRIRu
		
		Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Professor order by CO_Professor"	
		RSu.Open SQLu, CONu
		
while not RSu.eof
codigo_u = RSU("CO_Professor")
RSu.MOVENEXT
WEND
session("codigo_u") = codigo_u+1

case 1


		Set CONu = Server.CreateObject("ADODB.Connection") 
		ABRIRu = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONu.Open ABRIRu
		
		Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Usuario order by CO_Usuario"	
		RSu.Open SQLu, CONu
		
while not RSu.eof
codigo_u2 = RSU("CO_Usuario")
RSu.MOVENEXT
WEND
session("codigo_u2") = codigo_u2+1


end select
end function
%>
<%
' verifica se calcula média ou não

Function showmedia(curso,etapa,turma,co_materia)

if curso=2 then

if etapa=1 then
	select case co_materia
		case "CULJA"
			mostramedia = "mostra"
		case "CULJB"
			mostramedia = "mostra"
		case "EDFS"
			mostramedia = "mostra"	
		case else
			mostramedia = "nao"
	end select

elseif etapa=2 then
	select case co_materia
		case "CULJA"
			mostramedia = "mostra"
		case "CULJB"
			mostramedia = "mostra"
		case "EDFS"
			mostramedia = "mostra"	
		case else
			mostramedia = "nao"
	end select

elseif etapa=3 then
mostramedia = "mostra"
end if

elseif curso =1 and etapa=8 then
	select case co_materia
		case "EDFS"
			mostramedia = "mostra"
		case "HABA"
			mostramedia = "mostra"
		case "HEBR1"
			mostramedia = "mostra"
		case "HJUD2"
			mostramedia = "mostra"
		case "TANA2"
			mostramedia = "mostra"							
		case else
			mostramedia = "nao"
	end select
elseif curso =1 and etapa=88 then
	select case co_materia
		case "EDFS"
			mostramedia = "mostra"
		case "HABA"
			mostramedia = "mostra"
		case "HEBR2"
			mostramedia = "mostra"
		case "HJUD2"
			mostramedia = "mostra"
		case "TANA2"
			mostramedia = "mostra"							
		case else
			mostramedia = "nao"
	end select
end if
session("mostramedia")=mostramedia
end function

Function alterads(tipo,login_nv,pass_nv,mail_nv,cod,autorizo)
co_usr = cod
obr = request.QueryString("obr")

		Set conlg = Server.CreateObject("ADODB.Connection") 
		abrirlg = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		conlg.Open abrirlg
		
		Set conpf = Server.CreateObject("ADODB.Connection") 
		abrirpf = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		conpf.Open abrirpf

		Set RSlg = Server.CreateObject("ADODB.Recordset")
		SQLlg = "SELECT * FROM TB_Usuario WHERE CO_Usuario = "&co_usr 	
		RSlg.Open SQLlg, conlg

if RSlg.eof then
lg=""
sh=""
m1=""
aut=""
else
lg=RSlg("CO_Usuario")
sh=RSlg("Senha")	
ml=RSlg("TX_EMail_Usuario")
aut=RSlg("IN_Aut_email")
end if
Select case tipo
case 0
%>
<form action="index.asp?opt=cadastrar&obr=lg" method="post" name="cadastro" id="cadastro" onsubmit="return valid()">
        
  <table width="450" border="0" align="center" cellspacing="0">
    <tr> 
      <td width="115" class="form_tit_fundo"> <div align="right">Usu&aacute;rio atual :</div></td>
      <td><font class="form_dado_texto"> 
        <%  response.write(lg)%>
        </font> </td>
    </tr>
    <tr> 
      <td width="115" class="form_tit_fundo"> <div align="right">Usu&aacute;rio novo :</div></td>
      <td><input name="login" type="text" class="borda" id="login" size="50"> 
</td>
    </tr>
    <tr> 
      <td width="115"> <div align="right"></div></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td width="115">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="2"> <div align="center"> <font size="3" face="Courier New, Courier, mono">
          <input type="submit" name="Submit2" value="Confirmar" class="borda_bot2">
          </font></div></td>
    </tr>
  </table>
      </form>
  <% case 1
%>
<form action="index.asp?opt=cadastrar&obr=sh" method="post" name="cadastro" id="cadastro" onsubmit="return valid()">
            
        

  <table width="450" border="0" align="center" cellspacing="0">
    <tr> 
      <td width="115" class="form_tit_fundo"> <div align="right"><font class="style1"> Usu&aacute;rio :</font></div></td>
      <td><font class="style1"> 
        <%  response.write(lg)%>
        <input name="login" type="hidden" id="login" value="<%=lg%>">
        </font></td>
    </tr>
    <tr> 
      <td width="115" class="style1"> <div align="right">Senha :</div></td>
      <td><input name="pas1" type="password" id="pas1" class="borda"></td>
    </tr>
    <tr> 
      <td width="115" class="style1"> <div align="right">Confirma&ccedil;&atilde;o :</div></td>
      <td><input name="pas2" type="password" id="pas2" class="borda"></td>
    </tr>
    <tr> 
      <td width="115">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="2"> <div align="center"> <font size="3" face="Courier New, Courier, mono">
          <input type="submit" name="Submit3" value="Confirmar" class="borda_bot2">
          </font></div></td>
    </tr>
  </table>
          </form>
  <% case 2
%>
<form action="index.asp?opt=cadastrar&obr=ml" method="post" name="cadastro" id="cadastro" onsubmit="return checksubmit()">
            
        

  <table width="450" border="0" align="center" cellspacing="0">
    <tr> 
      <td width="115" class="form_tit_fundo"> <div align="right"><font class="style1"> 
          Usu&aacute;rio :</font></div></td>
      <td><font class="style1"> 
        <%  response.write(lg)%>
        <input name="login" type="hidden" id="login" value="<%=lg%>">
        </font></td>
    </tr>
    <tr> 
      <td width="115" class="form_tit_fundo"> <div align="right"><font class="style1"> 
          e-mail cadastrado:</font></div></td>
      <td><font class="style1"> 
        <%  response.write(ml)%>
        </font></td>
    </tr>
    <tr> 
      <td width="115" class="form_tit_fundo"> <div align="right"><font class="style1"> 
          novo e-mail :</font></div></td>
      <td><input name="email" type="text" class="borda" id="mail_nv" size="50"></td>
    </tr>
    <tr> 
      <td valign="top"> 
        <div align="right"> 
		<% if aut=TRUE then%>
          <input type="checkbox" name="autorizo" value="ok" checked/>
<%else%>
          <input type="checkbox" name="autorizo" value="ok" />
<%End if%>		  
        </div></td>
      <td><font class="style1">Autorizo o Web Fam&iacute;lia a enviar para o e-mail 
        informado <br>
        o usu&aacute;rio e a senha caso tenha esquecido.</font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="2"> <div align="center"> <font size="3" face="Courier New, Courier, mono"> 
          <input type="submit" name="Submit" value="Confirmar" class="borda_bot2">
          </font></div></td>
    </tr>
  </table>
          </form>		  
<%
case 99
if obr="lg" then
opcao="Login"
url="index.asp?opt=ok1"
log_tx="Login Alterado"

		Set RSlg = Server.CreateObject("ADODB.Recordset")
		SQLlg = "SELECT * FROM TB_Usuario WHERE Login = '"&login_nv& "'"	
		RSlg.Open SQLlg, conlg
if RSlg.eof then

		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "UPDATE TB_Usuario SET Login = '"&login_nv& "' WHERE CO_Usuario= " & co_usr
		RS.Open CONEXAO, conlg	

else
url="cadastro.asp?opt=err0"
end if

elseif obr="sh" then

opcao="Senha"
url="index.asp?opt=ok2"
log_tx="Senha Alterada"	
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "UPDATE TB_Usuario SET Senha = '"& pass_nv & "' WHERE CO_Usuario= " & co_usr
		RS.Open CONEXAO, conlg
		
elseif obr="ml" then
opcao="email"
url="index.asp?opt=ok3"
log_tx="E-mail Alterado"

		Set RSautorizo = Server.CreateObject("ADODB.Recordset")
		SQLautorizo = "SELECT * FROM TB_Usuario WHERE CO_Usuario= " & co_usr	
		RSautorizo.Open SQLautorizo, conlg
		
autorizo_anterior=RSautorizo("IN_Aut_email")

IF autorizo = "ok" then
autorizo= TRUE
ELSE
autorizo= FALSE
END IF


			ano = DatePart("yyyy", now)
			mes = DatePart("m", now) 
			dia = DatePart("d", now) 
			hora = DatePart("h", now) 
			min = DatePart("n", now) 

			data = dia &"/"& mes &"/"& ano

		Set RSlg = Server.CreateObject("ADODB.Recordset")
		SQLlg = "SELECT * FROM TB_Usuario WHERE TX_EMail_Usuario = '"&mail_nv& "'"	
'		SQLlg = "SELECT * FROM TB_Usuario WHERE CO_Usuario= " & co_usr	
		RSlg.Open SQLlg, conlg

if RSlg.eof then
		
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		CONEXAO2 = "UPDATE TB_Usuario SET TX_EMail_Usuario = '"&mail_nv& "', IN_Aut_email="& autorizo & ", DA_Cadastro='"&data&"' WHERE CO_Usuario= " & co_usr
		RS2.Open CONEXAO2, conlg
		
elseif autorizo_anterior<>autorizo then
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		CONEXAO2 = "UPDATE TB_Usuario SET TX_EMail_Usuario = '"&mail_nv& "', IN_Aut_email="& autorizo & ", DA_Cadastro='"&data&"' WHERE CO_Usuario= " & co_usr
		RS2.Open CONEXAO2, conlg

else
url="index.asp?opt=err1"
end if
		
end if		


			'call GravaLog ("WR-PR-PR-ALS",log_tx)		
		
response.Redirect(url)
End select
end function
%>