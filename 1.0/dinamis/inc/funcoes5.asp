<!--#include file="funcoes.asp"-->
<!--#include file="funcoes2.asp"-->
<!--#include file="funcoes6.asp"-->
<!--#include file="caminhos.asp"-->
<% 
Function alterads(tipo,login_nv,pass_nv,mail_nv,cod)
co_usr = cod

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
else
lg=RSlg("Login")
sh=RSlg("Senha")	
ml=RSlg("Email_Usuario")
end if
Select case tipo
case 0
%>
<form action="cadastro.asp?opt=cadastrar&obr=lg" method="post" name="cadastro" id="cadastro" onsubmit="return valid()">
        
  <table width="450" border="0" align="center" cellspacing="0">
    <tr> 
      <td width="115" class="tb_tit"> <div align="right">Usu&aacute;rio atual :</div></td>
      <td><font class="form_dado_texto"> 
        <%  response.write(lg)%>
        </font> </td>
    </tr>
    <tr> 
      <td width="115" class="tb_tit"> <div align="right">Usu&aacute;rio novo :</div></td>
      <td><input name="login" type="text" class="textInput" id="login" size="50"> 
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
      <td colspan="2"> <div align="center"> 
          <input type="submit" name="Submit" value="ok" class="botao_prosseguir">
        </div></td>
    </tr>
  </table>
      </form>
  <% case 1
%>
<form action="cadastro.asp?opt=cadastrar&obr=sh" method="post" name="cadastro" id="cadastro" onsubmit="return valid()">
            
        

  <table width="450" border="0" align="center" cellspacing="0">
    <tr> 
      <td width="115" class="tb_tit"> <div align="right">Usu&aacute;rio :</div></td>
      <td><font class="tb_tit"> 
        <%  response.write(lg)%>
        <input name="login" type="hidden" id="login" value="<%=lg%>">
        </font></td>
    </tr>
    <tr> 
      <td width="115" class="tb_tit"> <div align="right">Senha :</div></td>
      <td><input name="pas1" type="password" id="pas1" class="textInput"></td>
    </tr>
    <tr> 
      <td width="115" class="tb_tit"> <div align="right">Confirma&ccedil;&atilde;o :</div></td>
      <td><input name="pas2" type="password" id="pas2" class="textInput"></td>
    </tr>
    <tr> 
      <td width="115">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="2"> <div align="center"> 
          <input type="submit" name="Submit" value="ok" class="botao_prosseguir">
        </div></td>
    </tr>
  </table>
          </form>
  <% case 2
%>
<form action="cadastro.asp?opt=cadastrar&obr=ml" method="post" name="cadastro" id="cadastro" onsubmit="return valid()">
            
        

  <table width="450" border="0" align="center" cellspacing="0">
    <tr> 
      <td width="115" class="tb_tit"> <div align="right">Usu&aacute;rio :</div></td>
      <td><font class="form_dado_texto"> 
        <%  response.write(lg)%>
        <input name="login" type="hidden" id="login" value="<%=lg%>">
        </font></td>
    </tr>
    <tr> 
      <td width="115" class="tb_tit"> <div align="right">e-mail cadastrado :</div></td>
      <td><font class="form_dado_texto"> 
        <%  response.write(ml)%>
        </font></td>
    </tr>
    <tr> 
      <td width="115" class="tb_tit"> <div align="right">novo e-mail :</div></td>
      <td><input name="email" type="text" class="textInput" id="mail_nv" size="50"></td>
    </tr>
    <tr> 
      <td width="115">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="2"> <div align="center"> 
          <input type="submit" name="Submit" value="ok" class="botao_prosseguir">
        </div></td>
    </tr>
  </table>
          </form>		  
<%
case 99
if obr="lg" then
opcao="Login"
url="seguranca.asp?opt=ok1"
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
url="seguranca.asp?opt=ok2"
log_tx="Senha Alterada"	
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "UPDATE TB_Usuario SET Login = '"&login_nv& "', Senha = '"& pass_nv & "' WHERE CO_Usuario= " & co_usr
		RS.Open CONEXAO, conlg
		
elseif obr="ml" then
opcao="email"
url="seguranca.asp?opt=ok3"
log_tx="E-mail Alterado"

		Set RSlg = Server.CreateObject("ADODB.Recordset")
		SQLlg = "SELECT * FROM TB_Usuario WHERE Email_Usuario = '"&mail_nv& "'"	
		RSlg.Open SQLlg, conlg

if RSlg.eof then
		
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		CONEXAO2 = "UPDATE TB_Usuario SET Email_Usuario = '"&mail_nv& "' WHERE CO_Usuario= " & co_usr
		RS2.Open CONEXAO2, conlg
		
else
url="cadastro.asp?opt=err1"
end if
		
end if		


			call GravaLog ("WR-PR-PR-ALS",log_tx)		
		
response.Redirect(url)
End select
end function








'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function mapa_notas (CAMINHOa,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,opcao,erro,print)
P1=0
P2=0
P3=0
rec_ckeck="no"
res1=""
res2=""

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

select case mes
 case 1 
 mes = "janeiro"
 case 2 
 mes = "fevereiro"
 case 3 
 mes = "março"
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

if min<10 then
min = "0"&min
end if

data = dia &" de "& mes &" de "& ano
horario = hora & ":"& min

minimo_anual=70
minimo_final=50

		
			Set CON_N = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIR3
		
		Set CON_A = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_A.Open ABRIR
				
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidades &" AND CO_Curso = '"& grau &"' AND CO_Etapa = '"& serie &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)
		

select case opcao

case 1
tb="TB_NOTA_A"
case 2
tb="TB_NOTA_B"
case 3
tb="TB_NOTA_C"
end select
  %>
<table width="1000" border="0" align="right" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="17" class="tb_subtit"> <div align="right"><font class="form_dado_texto"> 
        N&ordm;</font></div></td>
    <td width="333" class="tb_subtit"> <div align="left"><font class="form_dado_texto"> 
        Nome</font></div></td>
    <td width="65" class="tb_subtit"> <div align="center"><font class="form_dado_texto"> 
        Média<BR>1º Tri</font></div></td>
    <td width="65" class="tb_subtit"> <div align="center"><font class="form_dado_texto"> 
        Média<BR>Acum.</font></div></td>
    <td width="65" class="tb_subtit"> <div align="center"><font class="form_dado_texto"> 
        Média<BR>2º Tri</font></div></td>
    <td width="65" class="tb_subtit"> <div align="center"><font class="form_dado_texto"> 
       Média<BR>Acum.</font></div></td>
    <td width="65" class="tb_subtit"> <div align="center"><font class="form_dado_texto"> 
        Média<BR>3º Tri</font></div></td>
    <td width="65" class="tb_subtit"> <div align="center"><font class="form_dado_texto"> 
        Média<BR>Acum.</font></div></td>
    <td width="65" class="tb_subtit"> <div align="center"><font class="form_dado_texto"> 
        RESULT</font></div></td>
    <td width="65" class="tb_subtit"> <div align="center"><font class="form_dado_texto"> 
        ECE</font></div></td>
    <td width="65" class="tb_subtit"> <div align="center"><font class="form_dado_texto"> 
        Média Final</font></div></td>
    <td width="65" class="tb_subtit"> <div align="center"><font class="form_dado_texto"> 
        RES.FINAL</font></div></td>
  </tr>
  <%check = 2
While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")

P1=0
P2=0
P3=0
	
  if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if 

  
  		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
		Set RS2 = CON_A.Execute(SQL_A)
	
NO_Aluno= RS2("NO_Aluno")
  
	  	Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"'"
		Set RS3 = CON_N.Execute(SQL_N)
	
if RS3.EOF then
%>
  <tr> 
    <td width="17" class="<%=cor%>"> <div align="center"><font class="form_dado_texto"> 
        &nbsp; </font></div></td>
    <td width="323" class="<%=cor%>"><font class="form_dado_texto">&nbsp;  </font></td>
    <td width="60" class="<%=cor%>"><div align="center"></div></td>
    <td width="60" class="<%=cor%>"><div align="center"></div></td>
    <td width="60" class="<%=cor%>"><div align="center"></div></td>
    <td width="60" class="<%=cor%>"><div align="center"></div></td>
    <td width="60" class="<%=cor%>"><div align="center"></div></td>
    <td width="60" class="<%=cor%>"><div align="center"></div></td>
    <td width="80" class="<%=cor%>"> <div align="center"><font class="form_dado_texto">  
        &nbsp; </font></div></td>
    <td width="80" class="<%=cor%>"> <div align="center"><font class="form_dado_texto">  
        &nbsp; </font></div></td>
    <td width="80" class="<%=cor%>"> <div align="center"><font class="form_dado_texto">  
        &nbsp; </font></div></td>
    <td width="80" class="<%=cor%>"> <div align="center"><font class="form_dado_texto">  
        &nbsp; </font></div></td>
  </tr>
  <%
else
		Me1= RS3("VA_Me1")
		Mc1= RS3("VA_Mc1")
		Me2= RS3("VA_Me2")
		Mc2= RS3("VA_Mc2")
		Me3= RS3("VA_Me3")
		Mc3= RS3("VA_Mc3")
if Mc3 >= minimo_anual then
res1="APR"
else
res1="ECE"
end if
		ece= RS3("VA_Me_EC")
		m3= RS3("VA_Mfinal")
		
if m3 >= minimo_final then
res2="APR"
else

res2="REP"
end if	

if isnull(Me1) or Me1="" then
else
Me1=Me1/10
Me1 = formatNumber(Me1,1) 
end if
if isnull(Mc1) or Mc1="" then
else
Mc1=Mc1/10
Mc1 = formatNumber(Mc1,1) 
end if
if isnull(Me2) or Me2="" then
else
Me2=Me2/10
Me2 = formatNumber(Me2,1) 
end if
if isnull(Mc2) or Mc2="" then
else
Mc2=Mc2/10
Mc2 = formatNumber(Mc2,1) 
end if
if isnull(Me3) or Me3="" then
else
Me3=Me3/10
Me3 = formatNumber(Me3,1) 
end if
if isnull(Mc3) or Mc3="" then
else
Mc3=Mc3/10
Mc3 = formatNumber(Mc3,1) 
end if
if isnull(ece) or ece="" then
else
ece=ece/10
ece = formatNumber(ece,1) 
end if
if isnull(m3) or m3="" then
else
m3=m3/10
m3 = formatNumber(m3,1) 
end if

if res1="APR" then
		ece= ""
		m3= ""
		res2=""
end if		
	%>
  <tr> 
    <td width="17" class="<%=cor%>"> <div align="center"><font class="form_dado_texto"> 
        <%response.Write(NU_Chamada)%>
        </font></div></td>
    <td width="323" class="<%=cor%>"><font class="form_dado_texto"> 
      <%response.Write(NO_Aluno)%>
      </font></td>
    <td width="60" class="<%=cor%>"><div align="center"><font class="form_dado_texto"> 
        <%response.Write(Me1)%>
        </font></div></td>
    <td width="60" class="<%=cor%>"><div align="center"><font class="form_dado_texto"> 
        <%response.Write(Mc1)%>
        </font></div></td>
    <td width="60" class="<%=cor%>"><div align="center"><font class="form_dado_texto"> 
        <%response.Write(Me2)%>
        </font></div></td>
    <td width="60" class="<%=cor%>"><div align="center"><font class="form_dado_texto"> 
        <%response.Write(Mc2)%>
        </font></div></td>
    <td width="60" class="<%=cor%>"><div align="center"><font class="form_dado_texto"> 
        <%response.Write(Me3)%>
        </font></div></td>
    <td width="60" class="<%=cor%>"><div align="center"><font class="form_dado_texto"> 
        <%response.Write(Mc3)%>
        </font></div></td>
    <td width="80" class="<%=cor%>"> <div align="center"><font class="form_dado_texto">  
        <%response.Write(res1)%>
        </font></div></td>
    <% end if


%>
    <td width="80" class="<%=cor%>"> <div align="center"><font class="form_dado_texto">  
        <%response.Write(ece)%>
        </font></div></td>
    <td width="80" class="<%=cor%>"> <div align="center"><font class="form_dado_texto">  
        <%response.Write(m3)%>
        </font></div></td>
    <td width="80" class="<%=cor%>"> <div align="center"><font class="form_dado_texto">  
        <%response.Write(res2)%>
        </font></div></td>
  </tr>
  <%
check = check+1
max=nu_chamada
RS.MoveNext
Wend 
session("max")=max
%>
</table>  
<%
End Function














'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function imp_mapa_notas (CAMINHOa,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,opcao,erro,print)

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

select case mes
 case 1 
 mes = "janeiro"
 case 2 
 mes = "fevereiro"
 case 3 
 mes = "março"
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

if min<10 then
min = "0"&min
end if

res1=""
res2=""

data = dia &" de "& mes &" de "& ano
horario = hora & ":"& min

minimo_anual=70
minimo_final=50



		
			Set CON_N = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIR3
		
		Set CON_A = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_A.Open ABRIR
				
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidades &" AND CO_Curso = '"& grau &"' AND CO_Etapa = '"& serie &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)
		

select case opcao

case 1
tb="TB_NOTA_A"
case 2
tb="TB_NOTA_B"
case 3
tb="TB_NOTA_C"
end select
  %>
<table width="1000" border="0" align="right" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="17" class="tabelaTit"> <div align="right"><font class="form_dado_texto"> 
        N&ordm;</font></div></td>
    <td width="323" class="tabelaTit"> <div align="left"><font class="form_dado_texto"> 
        Nome</font></div></td>
    <td width="60" class="tabelaTit"> <div align="center"><font class="form_dado_texto"> 
        Média 1º Tri</font></div></td>
    <td width="60" class="tabelaTit"> <div align="center"><font class="form_dado_texto"> 
        Média Acum.</font></div></td>
    <td width="60" class="tabelaTit"> <div align="center"><font class="form_dado_texto"> 
        Média 2º Tri</font></div></td>
    <td width="60" class="tabelaTit"> <div align="center"><font class="form_dado_texto"> 
       Média Acum.</font></div></td>
    <td width="60" class="tabelaTit"> <div align="center"><font class="form_dado_texto"> 
        Média 3º Tri</font></div></td>
    <td width="60" class="tabelaTit"> <div align="center"><font class="form_dado_texto"> 
        Média Acum.</font></div></td>
    <td width="60" class="tabelaTit"> <div align="center"><font class="form_dado_texto"> 
        RESULT</font></div></td>
    <td width="80" class="tabelaTit"> <div align="center"><font class="form_dado_texto"> 
        ECE</font></div></td>
    <td width="80" class="tabelaTit"> <div align="center"><font class="form_dado_texto"> 
        Média Final</font></div></td>
    <td width="80" class="tabelaTit"> <div align="center"><font class="form_dado_texto"> 
        RES.FINAL</font></div></td>
  </tr>
  <%check = 2
While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")

P1=0
P2=0
P3=0
	
  if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if 

  
  		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
		Set RS2 = CON_A.Execute(SQL_A)
	
NO_Aluno= RS2("NO_Aluno")
  
	  	Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"'"
		Set RS3 = CON_N.Execute(SQL_N)
	
if RS3.EOF then
%>
  <tr> 
    <td width="17" class="tabela"> <div align="center"><font class="form_dado_texto"> 
        &nbsp; </font></div></td>
    <td width="323" class="tabela"><font class="form_dado_texto">&nbsp;  </font></td>
    <td width="60" class="tabela"><div align="center"></div></td>
    <td width="60" class="tabela"><div align="center"></div></td>
    <td width="60" class="tabela"><div align="center"></div></td>
    <td width="60" class="tabela"><div align="center"></div></td>
    <td width="60" class="tabela"><div align="center"></div></td>
    <td width="60" class="tabela"><div align="center"></div></td>
    <td width="80" class="tabela"> <div align="center"><font class="form_dado_texto">  
        &nbsp; </font></div></td>
    <td width="80" class="tabela"> <div align="center"><font class="form_dado_texto">  
        &nbsp; </font></div></td>
    <td width="80" class="tabela"> <div align="center"><font class="form_dado_texto">  
        &nbsp; </font></div></td>
    <td width="80" class="tabela"> <div align="center"><font class="form_dado_texto">  
        &nbsp; </font></div></td>
  </tr>
  <%
else
		Me1= RS3("VA_Me1")
		Mc1= RS3("VA_Mc1")
		Me2= RS3("VA_Me2")
		Mc2= RS3("VA_Mc2")
		Me3= RS3("VA_Me3")
		Mc3= RS3("VA_Mc3")
if Mc3 >= minimo_anual then
res1="APR"
else
res1="ECE"
end if
		ece= RS3("VA_Me_EC")
		m3= RS3("VA_Mfinal")
		
if m3 >= minimo_final then
res2="APR"
else

res2="REP"
end if	

if isnull(Me1) or Me1="" then
else
Me1=Me1/10
Me1 = formatNumber(Me1,1) 
end if
if isnull(Mc1) or Mc1="" then
else
Mc1=Mc1/10
Mc1 = formatNumber(Mc1,1) 
end if
if isnull(Me2) or Me2="" then
else
Me2=Me2/10
Me2 = formatNumber(Me2,1) 
end if
if isnull(Mc2) or Mc2="" then
else
Mc2=Mc2/10
Mc2 = formatNumber(Mc2,1) 
end if
if isnull(Me3) or Me3="" then
else
Me3=Me3/10
Me3 = formatNumber(Me3,1) 
end if
if isnull(Mc3) or Mc3="" then
else
Mc3=Mc3/10
Mc3 = formatNumber(Mc3,1) 
end if
if isnull(ece) or ece="" then
else
ece=ece/10
ece = formatNumber(ece,1) 
end if
if isnull(m3) or m3="" then
else
m3=m3/10
m3 = formatNumber(m3,1) 
end if

if res1="APR" then
		ece= ""
		m3= ""
		res2=""
end if
	%>
  <tr> 
    <td width="17" class="tabela"> <div align="center"><font class="form_dado_texto"> 
        <%response.Write(NU_Chamada)%>&nbsp;
        </font></div></td>
    <td width="323" class="tabela"><font class="form_dado_texto"> 
      <%response.Write(NO_Aluno)%>&nbsp;
      </font></td>
    <td width="60" class="tabela"><div align="center"><font class="form_dado_texto"> 
        <%response.Write(Me1)%>&nbsp;
        </font></div></td>
    <td width="60" class="tabela"><div align="center"><font class="form_dado_texto"> 
        <%response.Write(Mc1)%>&nbsp;
        </font></div></td>
    <td width="60" class="tabela"><div align="center"><font class="form_dado_texto"> 
        <%response.Write(Me2)%>&nbsp;
        </font></div></td>
    <td width="60" class="tabela"><div align="center"><font class="form_dado_texto"> 
        <%response.Write(Mc2)%>&nbsp;
        </font></div></td>
    <td width="60" class="tabela"><div align="center"><font class="form_dado_texto"> 
        <%response.Write(Me3)%>&nbsp;
        </font></div></td>
    <td width="60" class="tabela"><div align="center"><font class="form_dado_texto"> 
        <%response.Write(Mc3)%>&nbsp;
        </font></div></td>
    <td width="80" class="tabela"> <div align="center"><font class="form_dado_texto">  
        <%response.Write(res1)%>&nbsp;
        </font></div></td>
    <% end if


%>
    <td width="80" class="tabela"> <div align="center"><font class="form_dado_texto">  
        <%response.Write(ece)%>&nbsp;
        </font></div></td>
    <td width="80" class="tabela"> <div align="center"><font class="form_dado_texto">  
        <%response.Write(m3)%>&nbsp;
        </font></div></td>
    <td width="80" class="tabela"> <div align="center"><font class="form_dado_texto">  
        <%response.Write(res2)%>&nbsp;
        </font></div></td>
  </tr>
  <%
check = check+1
max=nu_chamada
RS.MoveNext
Wend 
session("max")=max
%>
</table>
<%
End Function




'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function mapao (CAMINHOa,CAMINHOn,CAMINHO_pr,unidade,curso,co_etapa,turma,periodo,ano_letivo,co_usr,opcao,avaliacao,origem,erro)
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

select case mes
 case 1 
 mes = "janeiro"
 case 2 
 mes = "fevereiro"
 case 3 
 mes = "março"
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

if min<10 then
min = "0"&min
end if

data = dia &" de "& mes &" de "& ano
horario = hora & ":"& min

select case opcao

case 1
nota="TB_NOTA_A"
case 2
nota="TB_NOTA_B"
case 3
nota="TB_NOTA_C"
end select
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON3 = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3

		Set CON4 = Server.CreateObject("ADODB.Connection")
		ABRIR4 = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4

if origem = 0 then


		Set RSNN = Server.CreateObject("ADODB.Recordset")
		CONEXAONN = "Select CO_Materia from TB_Programa_Aula WHERE CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa&"' order by NU_Ordem_Boletim"
		Set RSNN = CON0.Execute(CONEXAONN)
		
materia_nome_check="vazio"
nome_nota="vazio"
i=0
largura = 0
While not RSNN.eof
materia_nome= RSNN("CO_Materia")

if materia_nome=materia_nome_check then
RS1.movenext
else

If Not IsArray(nome_nota) Then 
nome_nota = Array()
End if
If InStr(Join(nome_nota), materia_nome) = 0 Then
ReDim preserve nome_nota(UBound(nome_nota)+1)
nome_nota(Ubound(nome_nota)) = materia_nome
largura=largura+30

i=i+1
materia_nome_check=materia_nome

RSNN.movenext
end if
end if
wend
larg=770-(largura/i)

%>
<p>&nbsp;</p>
<p>&nbsp;</p>
<table width="770" border="0" align="right" cellspacing="0">
                <tr> 
                  <td width="17" bordercolor="#E9F0F8" bgcolor="#E9F0F8"> <div align="right"><font color="#0000CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>N&ordm;</strong></font></div></td>
                  <td width="larg" bordercolor="#E9F0F8" bgcolor="#E9F0F8"> <div align="center"><font color="#0000CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nome</strong></font></div></td>
                  <%For k=0 To ubound(nome_nota)%>
                  <td width="40" bordercolor="#E9F0F8" bgcolor="#E9F0F8"> <div align="center"><font color="#0000CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                      <% response.Write(nome_nota(k))%>
                      </strong></font></div></td>
                  <%
Next%>
                </tr>
                <%  check = 2
nu_chamada_check = 1

	Set RSA = Server.CreateObject("ADODB.Recordset")
	CONEXAOA = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
	Set RSA = CON4.Execute(CONEXAOA)
 
 While Not RSA.EOF
nu_matricula = RSA("CO_Matricula")
nu_chamada = RSA("NU_Chamada")

  		Set RSA2 = Server.CreateObject("ADODB.Recordset")
		CONEXAOA2 = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
		Set RSA2 = CON4.Execute(CONEXAOA2)
  		NO_Aluno= RSA2("NO_Aluno")

 if check mod 2 =0 then
  cor = "#F8FAFC" 
  else cor ="#F1F5FA"
  end if
  
if nu_chamada=nu_chamada_check then
nu_chamada_check=nu_chamada_check+1%>
                <tr bgcolor=<% = cor %>> 
                  <td width="17"> <div align="center"><font class="form_dado_texto"> <strong> 
                      <%response.Write(nu_chamada)%>
                      </strong></font></div></td>
                  <td width="200"> <div align="left"><font class="form_dado_texto"> <strong> 
                      <%response.Write(NO_Aluno)%>
                      </strong></font></div></td>
                  <%For k=0 To ubound(nome_nota)

  		Set RS3 = Server.CreateObject("ADODB.Recordset")
		CONEXAO3 = "Select VA_Media3 from "& nota & " WHERE NU_Periodo = "& periodo &" And CO_Materia = '"& nome_nota(k) &"' And CO_Matricula = "& nu_matricula
		Set RS3 = CON3.Execute(CONEXAO3)
  		
if RS3.EOF Then
%>
                  <td> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1"></font></font></div></td>
                  <%else
nota_materia= RS3("VA_Media3")%>
                  <td> <div align="center"><font class="form_dado_texto">  
                      <%response.Write(nota_materia)%>
                      </font></div></td>
                  <%end IF

NEXT%>
                </tr>
                <% 
else
While nu_chamada>nu_chamada_check
%>
                <tr bgcolor="#E4E4E4"> 
                  <td width="17" > <div align="center"><strong><font class="form_dado_texto">  
                      <%response.Write(nu_chamada_check)%>
                      </font></strong></div></td>
                  <td width="200"> <div align="left"><strong><font class="form_dado_texto">  
                      </font></strong></div></td>
                  <%For k=0 To ubound(nome_nota)%>
                  <td> <div align="center"></div></td>
                  <%

NEXT
%>
                </tr>
                <%
nu_chamada_check=nu_chamada_check+1	 
wend	
%>
                <tr bgcolor=<% = cor %>> 
                  <td width="17"> <div align="center"><font class="form_dado_texto"> <strong> 
                      <%response.Write(nu_chamada)%>
                      </strong></font></div></td>
                  <td width="200"> <div align="left"><font class="form_dado_texto"> <strong> 
                      <%response.Write(NO_Aluno)%>
                      </strong></font></div></td>
                  <%For k=0 To ubound(nome_nota)

  		Set RS3 = Server.CreateObject("ADODB.Recordset")
		CONEXAO3 = "Select VA_Media3 from "& nota & " WHERE NU_Periodo = "& periodo &" And CO_Materia = '"& nome_nota(k) &"' And CO_Matricula = "& nu_matricula
		Set RS3 = CON3.Execute(CONEXAO3)
  		
if RS3.EOF Then
%>
                  <td> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1"></font></font></div></td>
                  <%else
nota_materia= RS3("VA_Media3")%>
                  <td> <div align="center"><font size="2" face="Courier New, Courier, mono"> 
                      <%response.Write(nota_materia)%>
                      </font></div></td>
                  <%end IF

NEXT%>
                </tr>
                <%
 nu_chamada_check=nu_chamada_check+1	  
end if

	check = check+1
  RSA.MoveNext
  Wend 
%><tr><td colspan ="19">
<div align="Right">
<font class="form_dado_texto"> 
<% response.Write("impresso em "& data &" às "& horario)
%></font>
 </div>
</td></tr>
              </table>
			  
<%
elseif origem = 1 then
avaliacao = avaliacao
if nota="TB_NOTA_A" then%>
                        <%if avaliacao ="TES" then
campo_check ="VA_Teste" %>

                        <%end if%>
                        <%if avaliacao ="PRO" then
campo_check ="VA_Prova"%>

                        <%end if%>
                        <%if avaliacao ="N3" then
campo_check ="VA_3Nota"%>

                        <%end if%>
                        <%if avaliacao ="BON" then
campo_check ="VA_Bonus"%>

                        <%end if%>
                        <%if avaliacao ="REC" then
campo_check ="VA_Rec"%>

                        <%end if%>
                        <%elseif nota="TB_NOTA_B" then%>
                        <%if avaliacao ="A1" then
campo_check ="VA_Nota_A1"%>

                        <%end if%>
                        <%if avaliacao ="A2" then
campo_check ="VA_Nota_A2"%>

                        <%end if%>
                        <%if avaliacao ="B1" then
campo_check ="VA_Nota_B1"%>

                        <%end if%>
                        <%if avaliacao ="B2" then
campo_check ="VA_Nota_B2"%>

                        <%end if%>
                        <%if avaliacao ="AV1" then
campo_check ="VA_Nota1"%>

                        <%end if%>
                        <%if avaliacao ="AV2" then
campo_check ="VA_Nota2"%>

                        <%end if%>
                        <%if avaliacao ="N3" then
campo_check ="VA_Nota3"%>

                        <%end if%>
                        <%if avaliacao ="N4" then
campo_check ="VA_Nota4"%>

                        <%end if%>
                        <%if avaliacao ="BON" then
campo_check ="VA_Bonus"%>

                        <%end if%>
                        <%if avaliacao ="REC" then
campo_check ="VA_Rec"%>

                        <%end if%>
                        <%elseif nota="TB_NOTA_C" then%>
                        <%if avaliacao ="N1" then
campo_check ="VA_Nota1"%>

                        <%end if%>
                        <%if avaliacao ="N2" then
campo_check ="VA_Nota2"%>

                        <%end if%>
                        <%if avaliacao ="N3" then
campo_check ="VA_Nota3"%>

                        <%end if%>
                        <%if avaliacao ="BON" then
campo_check ="VA_Bonus"%>

                        <%end if%>
                        <%if avaliacao ="REC" then
campo_check ="VA_Rec"%>

                        <%end if%>
                        <%end if




		Set RSNN = Server.CreateObject("ADODB.Recordset")
		CONEXAONN = "Select CO_Materia from TB_Programa_Aula WHERE CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa&"' order by NU_Ordem_Boletim"
		Set RSNN = CON0.Execute(CONEXAONN)
		
materia_nome_check="vazio"
nome_nota="vazio"
i=0
largura = 0
While not RSNN.eof
materia_nome= RSNN("CO_Materia")

if materia_nome=materia_nome_check then
RS1.movenext
else

If Not IsArray(nome_nota) Then 
nome_nota = Array()
End if
If InStr(Join(nome_nota), materia_nome) = 0 Then
ReDim preserve nome_nota(UBound(nome_nota)+1)
nome_nota(Ubound(nome_nota)) = materia_nome
largura=largura+30

i=i+1
materia_nome_check=materia_nome

RSNN.movenext
end if
end if
wend
larg=770-(largura/i)

%>
              <table width="770" border="0" align="right" cellspacing="0">
                <tr> 
                  <td width="17" bordercolor="#E9F0F8" bgcolor="#E9F0F8"> <div align="right"><font color="#0000CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>N&ordm;</strong></font></div></td>
                  <td width="201" bordercolor="#E9F0F8" bgcolor="#E9F0F8"> <div align="center"><font color="#0000CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nome</strong></font></div></td>
                  <%For k=0 To ubound(nome_nota)%>
                  <td width="39" bordercolor="#E9F0F8" bgcolor="#E9F0F8"> <div align="center"><font color="#0000CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                      <% response.Write(nome_nota(k))%>
                      </strong></font></div></td>
                  <%
Next%>
                </tr>
                <%  check = 2
nu_chamada_check = 1

	Set RSA = Server.CreateObject("ADODB.Recordset")
	CONEXAOA = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
	Set RSA = CON4.Execute(CONEXAOA)
 
 While Not RSA.EOF
nu_matricula = RSA("CO_Matricula")
nu_chamada = RSA("NU_Chamada")

  		Set RSA2 = Server.CreateObject("ADODB.Recordset")
		CONEXAOA2 = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
		Set RSA2 = CON4.Execute(CONEXAOA2)
  		NO_Aluno= RSA2("NO_Aluno")

 if check mod 2 =0 then
  cor = "#F8FAFC" 
  else cor ="#F1F5FA"
  end if
  
if nu_chamada=nu_chamada_check then
nu_chamada_check=nu_chamada_check+1%>
                <tr bgcolor=<% = cor %>> 
                  <td width="17"> <div align="center"><font class="form_dado_texto"> <strong> 
                      <%response.Write(nu_chamada)%>
                      </strong></font></div></td>
                  <td width="201"> <div align="left"><font class="form_dado_texto"> <strong> 
                      <%response.Write(NO_Aluno)%>
                      </strong></font></div></td>
                  <%For k=0 To ubound(nome_nota)
 		Set RS3 = Server.CreateObject("ADODB.Recordset")
		CONEXAO3 = "Select "& campo_check & " from "& nota & " WHERE NU_Periodo = "& periodo &" And CO_Materia = '"& nome_nota(k) &"' And CO_Matricula = "& nu_matricula
		Set RS3 = CON3.Execute(CONEXAO3)
  		
if RS3.EOF Then
%>
                  <td> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1"></font></font></div></td>
                  <%else
nota_materia= RS3(""&campo_check&"")%>
                  <td width="505"> <div align="center"><font class="form_dado_texto">  
                      <%response.Write(nota_materia)%>
                      </font></div></td>
                  <%end IF

NEXT%>
                </tr>
                <% 
else
While nu_chamada>nu_chamada_check
%>
                <tr bgcolor="#E4E4E4"> 
                  <td width="17" > <div align="center"><strong><font class="form_dado_texto">  
                      <%response.Write(nu_chamada_check)%>
                      </font></strong></div></td>
                  <td width="201"> <div align="left"><strong><font class="form_dado_texto">  
                      </font></strong></div></td>
                  <%For k=0 To ubound(nome_nota)%>
                  <td> <div align="center"></div></td>
                  <%

NEXT
%>
                </tr>
                <%
nu_chamada_check=nu_chamada_check+1	 
wend	
%>
                <tr bgcolor=<% = cor %>> 
                  <td width="17"> <div align="center"><font class="form_dado_texto"> <strong> 
                      <%response.Write(nu_chamada)%>
                      </strong></font></div></td>
                  <td width="201"> <div align="left"><font class="form_dado_texto"> <strong> 
                      <%response.Write(NO_Aluno)%>
                      </strong></font></div></td>
                  <%For k=0 To ubound(nome_nota)

  		Set RS3 = Server.CreateObject("ADODB.Recordset")
		CONEXAO3 = "Select VA_Media3 from "& nota & " WHERE NU_Periodo = "& periodo &" And CO_Materia = '"& nome_nota(k) &"' And CO_Matricula = "& nu_matricula
		Set RS3 = CON3.Execute(CONEXAO3)
  		
if RS3.EOF Then
%>
                  <td> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1"></font></font></div></td>
                  <%else
nota_materia= RS3("VA_Media3")%>
                  <td> <div align="center"><font size="2" face="Courier New, Courier, mono"> 
                      <%response.Write(nota_materia)%>
                      </font></div></td>
                  <%end IF

NEXT%>
                </tr>
                <%
 nu_chamada_check=nu_chamada_check+1	  
end if

	check = check+1
  RSA.MoveNext
  Wend 
%><tr><td colspan ="19">
<div align="Right">
<font class="form_dado_texto"> 
<% response.Write("impresso em "& data &" às "& horario)
%></font>
 </div>
</td></tr>
              </table>
<%
end if 
end function			  
%>

