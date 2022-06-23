<!--#include file="caminhos.asp"-->


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

'///////////////////////////////////////////////    decode    //////////////////////////////////////////////////////////////////////////////
Function DecodificaServerUrl(nome_a_alterar)
str = Replace(nome_a_alterar, "+", " ") 
        For n = 1 To Len(str) 
            sT = Mid(str, n, 1) 
            If sT = "%" Then 
                If n+2 < Len(str) Then 
                    sR = sR & _ 
                        Chr(CLng("&H" & Mid(str, n+1, 2))) 
                    n = n+2 
                End If 
            Else 
                sR = sR & sT 
            End If 
        Next 
        DecodificaServerUrl = sR
End Function
Function GeraNomesNovaVersao(tipo_dado,variavel1,variavel2,variavel3,variavel4,variavel5,conexao,outro)

	if tipo_dado="Mun" then
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Municipios where SG_UF ='"& variavel1 &"' AND CO_Municipio = "&variavel2
		RS.Open SQL, conexao	
	
		IF RS.eof THEN
			GeraNomesNovaVersao= ""
		ELSE
			GeraNomesNovaVersao= RS("NO_Municipio")
		END IF
	elseif tipo_dado="Bai" then
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Bairros where CO_Bairro ="& variavel3 &"AND SG_UF ='"& variavel1&"' AND CO_Municipio = "&variavel2
		RS.Open SQL, conexao

		IF RS.eof THEN
			GeraNomesNovaVersao= ""
		ELSE
			GeraNomesNovaVersao= RS("NO_Bairro")
		END IF		
	elseif tipo_dado="D" then
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Materia where CO_Materia = '"& variavel1&"'"	
		RS.Open SQL, conexao
	
		IF RS.eof THEN
			GeraNomesNovaVersao= ""
		ELSE
			GeraNomesNovaVersao= RS("NO_Materia")
		END IF
		
	elseif tipo_dado="U" then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Unidade where NU_Unidade = "& variavel1
		RS.Open SQL, conexao
	
		IF RS.eof THEN
			GeraNomesNovaVersao= ""
		ELSE
			GeraNomesNovaVersao= RS("NO_Unidade")
		END IF
		
	elseif tipo_dado="IU" then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Unidade where NU_Unidade = "& variavel1
		RS.Open SQL, conexao
	
		IF RS.eof THEN
			GeraNomesNovaVersao= ""
		ELSE
			GeraNomesNovaVersao= RS("TX_Imp_Cabecalho")
		END IF		

	elseif tipo_dado="C" then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Curso where CO_Curso = '"& variavel1 &"'"
		RS.Open SQL, conexao
	
		IF RS.eof THEN
			GeraNomesNovaVersao= ""
		ELSE
			GeraNomesNovaVersao= RS("NO_Curso")
		END IF
	elseif tipo_dado="CA" then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Curso where CO_Curso = '"& variavel1 &"'"
		RS.Open SQL, conexao
	
		IF RS.eof THEN
			GeraNomesNovaVersao= ""
		ELSE
			GeraNomesNovaVersao= RS("NO_Abreviado_Curso")
		END IF			
	elseif tipo_dado="PC" then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Curso where CO_Curso = '"& variavel1 &"'"
		RS.Open SQL, conexao
	
		IF RS.eof THEN
			GeraNomesNovaVersao= ""
		ELSE
			GeraNomesNovaVersao= RS("CO_Conc")
		END IF	
			
	elseif tipo_dado="E" then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Etapa where CO_Curso = '"& variavel1 &"' and CO_Etapa = '"& variavel2 &"'"
		RS.Open SQL, conexao
	
		IF RS.eof THEN
			GeraNomesNovaVersao= ""
		ELSE
			GeraNomesNovaVersao= RS("NO_Etapa")
		END IF	
		
	elseif tipo_dado="SA" then	
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Situacao_Aluno where CO_Situacao = '"& variavel1 &"'"
		RS.Open SQL, conexao
	
		IF RS.eof THEN
			GeraNomesNovaVersao= ""
		ELSE
			GeraNomesNovaVersao= RS("TX_Descricao_Situacao")
		END IF		
		
	elseif tipo_dado="GRP_ITEM" then	
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Grupo where CO_Grupo = "& variavel1
		RS.Open SQL, conexao
	
		IF RS.eof THEN
			GeraNomesNovaVersao= ""
		ELSE
			GeraNomesNovaVersao= RS("NO_Grupo")
		END IF					
		
elseif tipo_dado="MES_ABR" then
	
		variavel1=variavel1*1	
		IF variavel1=1 THEN
			GeraNomesNovaVersao= "Jan"
		ELSEIF variavel1=2 THEN
			GeraNomesNovaVersao= "Fev"
		ELSEIF variavel1=3 THEN
			GeraNomesNovaVersao= "Mar"
		ELSEIF variavel1=4 THEN
			GeraNomesNovaVersao= "Abr"
		ELSEIF variavel1=5 THEN
			GeraNomesNovaVersao= "Mai"
		ELSEIF variavel1=6 THEN
			GeraNomesNovaVersao= "Jun"
		ELSEIF variavel1=7 THEN
			GeraNomesNovaVersao= "Jul"
		ELSEIF variavel1=8 THEN
			GeraNomesNovaVersao= "Ago"
		ELSEIF variavel1=9 THEN
			GeraNomesNovaVersao= "Set"
		ELSEIF variavel1=10 THEN
			GeraNomesNovaVersao= "Out"
		ELSEIF variavel1=11 THEN
			GeraNomesNovaVersao= "Nov"
		ELSEIF variavel1=12 THEN
			GeraNomesNovaVersao= "Dez"																														
		END IF	

	elseif tipo_dado="MES" then
	
		variavel1=variavel1*1	
		IF variavel1=1 THEN
			GeraNomesNovaVersao= "Janeiro"
		ELSEIF variavel1=2 THEN
			GeraNomesNovaVersao= "Fevereiro"
		ELSEIF variavel1=3 THEN
			GeraNomesNovaVersao= "Mar&ccedil;o"
		ELSEIF variavel1=4 THEN
			GeraNomesNovaVersao= "Abril"
		ELSEIF variavel1=5 THEN
			GeraNomesNovaVersao= "Maio"
		ELSEIF variavel1=6 THEN
			GeraNomesNovaVersao= "Junho"
		ELSEIF variavel1=7 THEN
			GeraNomesNovaVersao= "Julho"
		ELSEIF variavel1=8 THEN
			GeraNomesNovaVersao= "Agosto"
		ELSEIF variavel1=9 THEN
			GeraNomesNovaVersao= "Setembro"
		ELSEIF variavel1=10 THEN
			GeraNomesNovaVersao= "Outubro"
		ELSEIF variavel1=11 THEN
			GeraNomesNovaVersao= "Novembro"
		ELSEIF variavel1=12 THEN
			GeraNomesNovaVersao= "Dezembro"																														
		END IF					
								
		
	END IF	

end Function




Function GeraNomes(tipo_dado,variavel1,variavel2,variavel3,variavel4,variavel5,conexao,outro)

	if tipo_dado="Mun" then
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Municipios where SG_UF ='"& variavel1 &"' AND CO_Municipio = "&variavel2
		RS.Open SQL, conexao	
	
		IF RS.eof THEN
			GeraNomes= ""
		ELSE
			GeraNomes= RS("NO_Municipio")
		END IF
	elseif tipo_dado="Bai" then
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Bairros where CO_Bairro ="& variavel3 &"AND SG_UF ='"& variavel1&"' AND CO_Municipio = "&variavel2
		RS.Open SQL, conexao

		IF RS.eof THEN
			GeraNomes= ""
		ELSE
			GeraNomes= RS("NO_Bairro")
		END IF		
	elseif tipo_dado="D" then
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Materia where CO_Materia = '"& variavel1&"'"	
		RS.Open SQL, conexao
	
		IF RS.eof THEN
			GeraNomes= ""
		ELSE
			GeraNomes= RS("NO_Materia")
		END IF
		
	elseif tipo_dado="U" then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Unidade where NU_Unidade = "& variavel1
		RS.Open SQL, conexao
	
		IF RS.eof THEN
			GeraNomes= ""
		ELSE
			GeraNomes= RS("NO_Unidade")
		END IF

	elseif tipo_dado="C" then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Curso where CO_Curso = '"& variavel1 &"'"
		RS.Open SQL, conexao
	
		IF RS.eof THEN
			GeraNomes= ""
		ELSE
			GeraNomes= RS("NO_Curso")
		END IF
	elseif tipo_dado="PC" then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Curso where CO_Curso = '"& variavel1 &"'"
		RS.Open SQL, conexao
	
		IF RS.eof THEN
			GeraNomes= ""
		ELSE
			GeraNomes= RS("CO_Conc")
		END IF	
			
	elseif tipo_dado="E" then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Etapa where CO_Curso = '"& variavel1 &"' and CO_Etapa = '"& variavel2 &"'"
		RS.Open SQL, conexao
	
		IF RS.eof THEN
			GeraNomes= ""
		ELSE
			GeraNomes= RS("NO_Etapa")
		END IF	
		
	elseif tipo_dado="SA" then	
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Situacao_Aluno where CO_Situacao = '"& variavel1 &"'"
		RS.Open SQL, conexao
	
		IF RS.eof THEN
			GeraNomes= ""
		ELSE
			GeraNomes= RS("TX_Descricao_Situacao")
		END IF			
		
	END IF

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
          <input type="submit" name="Submit2" value=" " class="confirmar">
          </font></div></td>
    </tr>
  </table>
      </form>
  <% case 1
%>
<form action="index.asp?opt=cadastrar&obr=sh" method="post" name="cadastro" id="cadastro" onsubmit="return valid()">
            
        

  <table width="450" border="0" align="center" cellspacing="0">
    <tr> 
      <td width="115" class="form_tit_fundo"> <div align="right"><font class="style1"> 
          Usu&aacute;rio :</font></div></td>
      <td colspan="2"><font class="style1"> 
        <%  response.write(lg)%>
        <input name="login" type="hidden" id="login" value="<%=lg%>">
        </font></td>
    </tr>
    <tr> 
      <td width="115" class="style1"> <div align="right">Senha :</div></td>
      <td colspan="2"><input name="pas1" type="password" id="pas1" class="borda"></td>
    </tr>
    <tr> 
      <td width="115" class="style1"> <div align="right">Confirma&ccedil;&atilde;o 
          :</div></td>
      <td colspan="2"><input name="pas2" type="password" id="pas2" class="borda"></td>
    </tr>
    <tr> 
      <td width="115">&nbsp;</td>
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> <div align="center"> <font size="3" face="Courier New, Courier, mono"> 
          </font></div></td>
      <td width="140"><div align="center"><font size="3" face="Courier New, Courier, mono"> 
          <input type="submit" name="Submit3" value=" " class="confirmar">
          </font></div></td>
      <td width="189">&nbsp;</td>
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
          <input type="submit" name="Submit" value=" " class="confirmar">
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

Function regra_aprovacao (curso,etapa,m1_aluno,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,tipo_calculo)

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set RSra = Server.CreateObject("ADODB.Recordset")
	SQLra = "SELECT * FROM TB_Regras_Aprovacao where CO_Curso = '"&curso&"' and CO_Etapa = '"&etapa&"'"
	RSra.Open SQLra, CON0	
			
	valor_m1=RSra("NU_Valor_M1")
	m1_menor=RSra("NU_Int_Me_Ma_Igual_M1")
	m1_maior_igual=RSra("NU_Int_Me_Me_M1")
	res1_3=RSra("NO_Expr_Ma_Igual_M1")
	res1_2=RSra("NO_Expr_Int_M1_V")
	res1_1=RSra("NO_Expr_Int_M1_F")
	peso_m2_m1=RSra("NU_Peso_Media_M2_M1")
	peso_m2_m2=RSra("NU_Peso_Media_M2_M2")
	
	valor_m2=RSra("NU_Valor_M2")
	m2_menor=RSra("NU_Int_Me_Ma_Igual_M2")
	m2_maior_igual=RSra("NU_Int_Me_Me_M2")	
	res2_3=RSra("NO_Expr_Ma_Igual_M2")
	res2_2=RSra("NO_Expr_Int_M2_V")
	res2_1=RSra("NO_Expr_Int_M2_F")
	peso_m3_m1=RSra("NU_Peso_Media_M3_M1")
	peso_m3_m2=RSra("NU_Peso_Media_M3_M2")
	peso_m3_m3=RSra("NU_Peso_Media_M3_M3")
	
	valor_m3=RSra("NU_Valor_M3")
	m3_menor=RSra("NU_Int_Me_Ma_Igual_M3")
	m3_maior_igual=RSra("NU_Int_Me_Me_M3")	
	res3_1=RSra("NO_Expr_Int_M3_F")
	res3_2=RSra("NO_Expr_Int_M3_V")

		
	m1_aluno=m1_aluno*1	
	m1_maior_igual=m1_maior_igual*1
	m1_menor=m1_menor*1


	if m1_aluno >= m1_maior_igual then
		resultado=res1_3
		resultado1="apr"
	elseif m1_aluno >= m1_menor then
		resultado=res1_2
	else
		resultado=res1_1	
	end if
	
	if tipo_calculo="wfboletim" then
		m1_wfboletim=m1_aluno
		resultado1_wfboletim=resultado
	end if		
'response.Write("if "&m1_aluno &">="& m1_maior_igual &"then<BR>")
'response.Write("elseif "&m1_aluno &">"& m1_menor &"then<BR>")	
'response.Write(resultado&"<BR>")	
	if resultado1="apr" then
		m2_aluno=m1_aluno	
		m3_aluno=m1_aluno
		if tipo_calculo="wfboletim" then
			m2_wfboletim="&nbsp;"
			m3_wfboletim="&nbsp;"			
			resultado2_wfboletim="&nbsp;"
			resultado3_wfboletim="&nbsp;"		
		end if			
	else			
		if tipo_calculo="recuperacao" or tipo_calculo="final" or tipo_calculo="wfboletim" then
			if nota_aux_m2_1="&nbsp;" then
				m2_aluno="&nbsp;"
				resultado="&nbsp;"					
			else								
				m1_aluno_peso=m1_aluno*peso_m2_m1
				nota_aux_m2_1_peso=nota_aux_m2_1*peso_m2_m2
				m2_aluno=(m1_aluno_peso+nota_aux_m2_1_peso)/(peso_m2_m1+peso_m2_m2)
				m2_aluno=m2_aluno*10
				decimo = m2_aluno - Int(m2_aluno)
				If decimo >= 0.5 Then
					nota_arredondada = Int(m2_aluno) + 1
					m2_aluno=nota_arredondada
				else
					nota_arredondada = Int(m2_aluno)
					m2_aluno=nota_arredondada											
				End If	
				m2_aluno=m2_aluno/10				
				m2_aluno = formatNumber(m2_aluno,1)
				m2_aluno=m2_aluno*1
				m2_maior_igual=m2_maior_igual*1	
				m2_menor=m2_menor*1		
				if m2_aluno >= m2_maior_igual then
					resultado=res2_3
					resultado2="apr"
				elseif m2_aluno >= m2_menor then
					resultado=res2_2
				else
					resultado=res2_1	
				end if
				
				if tipo_calculo="wfboletim" then
					m2_wfboletim=m2_aluno	
					resultado2_wfboletim=resultado	
					m3_wfboletim="&nbsp;"			
					resultado3_wfboletim="&nbsp;"						
				end if	

			end if
 			if	tipo_calculo="final" or tipo_calculo="wfboletim" then
				if resultado2="apr" then
					m3_aluno=m2_aluno									
				else
					if m2_aluno="&nbsp;" or nota_aux_m2_1="&nbsp;" or nota_aux_m3_1="&nbsp;" then		
						m3_aluno="&nbsp;"
						resultado="&nbsp;"			
					else								
						m1_aluno_peso=m1_aluno*peso_m3_m1
						m2_aluno_peso=m2_aluno*peso_m3_m2
						nota_aux_m3_1_peso=nota_aux_m3_1*peso_m3_m3
						m3_aluno=(m1_aluno_peso+m2_aluno_peso+nota_aux_m3_1_peso)/(peso_m3_m1+peso_m3_m2+peso_m3_m3)
						m3_aluno=m3_aluno*10
						decimo = m3_aluno - Int(m3_aluno)
						If decimo >= 0.5 Then
							nota_arredondada = Int(m3_aluno) + 1
							m3_aluno=nota_arredondada
						else
							nota_arredondada = Int(m3_aluno)
							m3_aluno=nota_arredondada											
						End If	
						m3_aluno=m3_aluno/10
						m3_aluno = formatNumber(m3_aluno,1)
						m3_aluno=m3_aluno*1
						valor_m3=valor_m3*1		
						m3_maior_igual=m3_maior_igual*1		
						if m3_aluno >= m3_maior_igual then
							resultado=res3_2
						else
							resultado=res3_1	
						end if					
					end if
					if tipo_calculo="wfboletim" then
						m3_wfboletim=m3_aluno		
						resultado3_wfaboletim=resultado	
					end if						
				end if
			end if	
		end if	
	end if

	if tipo_calculo="anual" then
		m1_aluno = formatNumber(m1_aluno,1)	
		regra_aprovacao=m1_aluno&"#!#"&resultado
	elseif tipo_calculo="recuperacao" then
		if resultado1="apr" then
			m1_aluno = formatNumber(m1_aluno,1)	
			regra_aprovacao=m1_aluno&"#!#"&resultado		
		else
			if m2_aluno<>"&nbsp;" then
				m2_aluno = formatNumber(m2_aluno,1)
			end if
			regra_aprovacao=m2_aluno&"#!#"&resultado
		end if
	elseif tipo_calculo="wfboletim" then
				if m2_aluno<>"&nbsp;" then
					m2_aluno = formatNumber(m2_aluno,0)
				end if	
	
				if m3_aluno<>"&nbsp;" then
					m3_aluno = formatNumber(m3_aluno,0)
				end if
			regra_aprovacao=m1_wfboletim&"#!#"&resultado1_wfboletim&"#!#"&m2_wfboletim&"#!#"&resultado2_wfboletim&"#!#"&m3_wfboletim&"#!#"&resultado3_wfboletim
	else
		if resultado2="apr"then
			if m2_aluno<>"&nbsp;" then
				m2_aluno = formatNumber(m2_aluno,1)
			end if
			regra_aprovacao=m2_aluno&"#!#"&resultado		
		else
			if m3_aluno<>"&nbsp;" then
				m3_aluno = formatNumber(m3_aluno,1)
			end if
			regra_aprovacao=m3_aluno&"#!#"&resultado			
		end if
	end if
	
	'Session("M2")=m2_aluno
	'Session("M3")=m3_aluno
end function		

'///////////////////////////////////////////////    VETOR     //////////////////////////////////////////////////////////////////////////////


Function VetorMonta(Acao,Valor)
'Usamos o case para manipular a ação da função
Select Case Trim(Acao)
'Inclui nova posicao ao vetor
Case "Incluir"
'Guarda na variavel Vetor o conteudo da Session
Vetor = Session("GuardaVetor")
'Verifica se a Variavel Vetor é um Array, caso nao for entao definimos ela um Array
If Not IsArray(Vetor) Then Vetor = Array() End if
'Verifica se o Valor que esta sendo inserido já esta no Vetor se estiver entao nao inseri para nao haver duplicidades do vetor
If InStr(Join(Vetor), Valor) = 0 Then
'Este comando ira preservar o vetor e adciona + 1 valor
ReDim preserve Vetor(UBound(Vetor)+1)
'Este é o valor que estamos adicionando no vetor
Vetor(Ubound(Vetor )) = Valor
'Coloca o conteudo da variavel vetor dentro da Session
Session("GuardaVetor") = Vetor
End if
'Apaga uma determinada posicao do vetor
Case "Excluir"
'Inicia a varivel vetor como vazia
Vetor = Array()
'Criamos uma nova variavel Auxiliar e guardamos o valor da Session
AuxVetor = Session("GuardaVetor")
'Definine a Session como um Array vazio
Session("GuardaVetor") = Array()
'Faz um laço em todas as posições do vetor
For i = 0 To Ubound(AuxVetor)
'Verifica se o valor passado para excluir do vetor é diferente do valor que esta dentro da variavel Auxiliar
If AuxVetor(i) <> (Valor) Then
'Este comando ira preservar o vetor e adciona + 1 valor
ReDim preserve Vetor (UBound(Vetor)+1)
'Este é o valor que estamos adicionando no vetor
Vetor (Ubound(Vetor)) = AuxVetor(i)
'Coloca o conteudo da variavel vetor dentro da Session
Session("GuardaVetor") = Vetor
End If
Next
'Fim do Case
End Select
End Function

Function Incluir_Vetor

'Executa a função que ira criar uma posição do vetor, basta passar a acao e o valor
Call VetorMonta("Incluir",Valor_Vetor)
'Request("Valor_Vetor")
'response.Write(Valor_Vetor&"=vet<BR>")
End Function


Function VisualizaValoresVetor
vet = session("GuardaVetor")

'Veriofica se a Session é um array, caso nao for então atribuimos a Session como um Array
IF Not IsArray(vet) Then vet = Array() End if
'Faremos um laço entre todos os vetores criados

if ubound(vet) >0 then
%>
<table width="1000" border="0" cellspacing="0">
  <tr> 
    <td valign="top"> 
      <table width="100%" border="0" align="right" cellspacing="0">
        <tr class="tb_corpo"
> 
          <td class="tb_tit"
>Lista de Professores</td>
        </tr>
        <tr> 
          <td> <ul>
              <%
For x = 0 To ubound(vet) 

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Professor where NO_Professor = '"&vet(x)&"' order BY NO_Professor"
		RS.Open SQL, CON1


cod_cons = RS("CO_Professor")
ativo = RS("IN_Ativo_Escola")
if ativo = "True" then
Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=ativos href=altera.asp?ch=ok&ori=01&cod_cons="&cod_cons&"&nvg="&nvg&" >"&vet(x)&"</a></font></li>")
else
Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=inativos href=altera.asp?ch=ok&ori=01&cod_cons="&cod_cons&"&nvg="&nvg&">"&vet(x)&"</a></font></li>")
end if
'Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a href =altera.asp?or=02&cod="&cod&">"&vet(x)&"</a></font></li>")
Next
%>
            </ul></td>
        </tr>
      </table></td>
  </tr>
</table>
<%

elseif ubound(vet)=0 then

		Set RS = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT * FROM TB_Professor where NO_Professor like '%"& strProcura & "%' order BY NO_Professor"
		RS.Open SQL, CON1


cod_cons = RS("CO_Professor")

response.Redirect("altera.asp?ori=01&cod_cons="&cod_cons&"&nvg="&nvg)
else

response.Redirect("index.asp?ori=01&opt=err2&cod_cons="&cod_cons&"&nvg="&nvg)%>

<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<%
end if
'Verifica se a Session tem alguma posição, se tiver mostra a opção de apagar todos os vetores
'If ubound(vet) >= 0 Then
'Response.Write "<br>" &"<a href='vetor.asp?action=LimpaVetor'>Apagar Tudo</a>" & "<br>" 'Imprime o Vetor na tela
'End if

End Function

Function LimpaVetor

'Limpa todas as posiçoes do vetor, apagando a Session
session("GuardaVetor") = Empty

End Function
'///////////////////////////////////////// vetor alunos /////////////////////////////////////////////////////////////////
Function VetorMonta2(Acao,Valor)
'Usamos o case para manipular a ação da função
Select Case Trim(Acao)
'Inclui nova posicao ao vetor
Case "Incluir"
'Guarda na variavel Vetor o conteudo da Session
Vetor = Session("GuardaVetor")
'Verifica se a Variavel Vetor é um Array, caso nao for entao definimos ela um Array
If Not IsArray(Vetor) Then Vetor = Array() End if
'Verifica se o Valor que esta sendo inserido já esta no Vetor se estiver entao nao inseri para nao haver duplicidades do vetor
If InStr(Join(Vetor), Valor) = 0 Then
'Este comando ira preservar o vetor e adciona + 1 valor
ReDim preserve Vetor(UBound(Vetor)+1)
'Este é o valor que estamos adicionando no vetor
Vetor(Ubound(Vetor )) = Valor
'Coloca o conteudo da variavel vetor dentro da Session
Session("GuardaVetor") = Vetor
End if
'Apaga uma determinada posicao do vetor
Case "Excluir"
'Inicia a varivel vetor como vazia
Vetor = Array()
'Criamos uma nova variavel Auxiliar e guardamos o valor da Session
AuxVetor = Session("GuardaVetor")
'Definine a Session como um Array vazio
Session("GuardaVetor") = Array()
'Faz um laço em todas as posições do vetor
For i = 0 To Ubound(AuxVetor)
'Verifica se o valor passado para excluir do vetor é diferente do valor que esta dentro da variavel Auxiliar
If AuxVetor(i) <> (Valor) Then
'Este comando ira preservar o vetor e adciona + 1 valor
ReDim preserve Vetor (UBound(Vetor)+1)
'Este é o valor que estamos adicionando no vetor
Vetor (Ubound(Vetor)) = AuxVetor(i)
'Coloca o conteudo da variavel vetor dentro da Session
Session("GuardaVetor") = Vetor
End If
Next
'Fim do Case
End Select
End Function

Function Incluir_Vetor2

'Executa a função que ira criar uma posição do vetor, basta passar a acao e o valor
Call VetorMonta("Incluir",Valor_Vetor)
'Request("Valor_Vetor")
'response.Write(Valor_Vetor&"=vet<BR>")
End Function


Function VisualizaValoresVetor2
vet = session("GuardaVetor")

'Veriofica se a Session é um array, caso nao for então atribuimos a Session como um Array
IF Not IsArray(vet) Then vet = Array() End if
'Faremos um laço entre todos os vetores criados

if ubound(vet) >0 then
%>
  <tr> 
    <td valign="top"> 
<table width="1000" border="0" cellspacing="0">
  <tr> 
    <td valign="top"> 
      
      <table width="100%" border="0" align="right" cellspacing="0">
        <tr class="tb_corpo"
> 
          <td class="tb_tit"
>Lista de Alunos</td>
        </tr>
        <tr> 
          <td> <ul>
              <%
For x = 0 To ubound(vet) 


		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos where CO_Matricula = "&vet(x)&" order BY NO_Aluno"
		RS.Open SQL, CON1

cod_cons =vet(x) 
no_aluno = RS("NO_Aluno")
Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=ativos href =altera.asp?ori=01&cod_cons="&cod_cons&"&nvg="&nvg&">"&no_aluno&"</a></font></li>")
Next
%>
            </ul></td>
        </tr>
      </table></td>
  </tr>
</table>
<%

elseif ubound(vet)=0 then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos where NO_Aluno like '%"& strProcura & "%' order BY NO_Aluno"
		RS.Open SQL, CON1


cod_cons = RS("CO_Matricula")

response.Redirect("altera.asp?ori=01&cod_cons="&cod_cons&"&nvg="&nvg)
else
Session("nome_cadastrar")=strProcura
%>
            <tr> 
              
    <td height="10" colspan="5"> 
      <%call mensagens(nivel,mensagem,1,0) %>
    </td>
			   </tr>
            <tr> 
              
    <td height="10" colspan="5"> 
      <%call mensagens(nivel,300,0,0) %>
    </td>
			  </tr>

        <tr> 
            <td valign="top"> 			  
        <form action="index.asp?opt=list&nvg=<%=nvg%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
          <table width="1000" border="0" cellspacing="0">
            <tr> 		
                  <tr class="tb_tit"> 
                    
      <td height="10" colspan="5">Preencha um dos campos abaixo</td>
                  </tr>
                  <tr> 
                    
      <td width="10"  height="10"> 
        <div align="right"><font class="form_dado_texto"> Matr&iacute;cula: 
          </font></div></td>
                    
      <td width="10"  height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
        </font><font size="2" face="Arial, Helvetica, sans-serif"> 
        <input name="busca1" type="text" class="textInput" id="busca1" size="12">
                      </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
                      </font></td>
                    
      <td width="10" height="10"> 
        <div align="right"><font class="form_dado_texto"> Nome: 
                        </font></div></td>
                    
      <td width="10"  height="10" ><font size="2" face="Arial, Helvetica, sans-serif"> 
        <input name="busca2" type="text" class="textInput" id="busca2" size="55" maxlength="50">
                      </font></td>
                    
      <td width="10" height="10" ><font size="2" face="Arial, Helvetica, sans-serif"> 
        <input name="Submit3" type="submit" class="botao_prosseguir" id="Submit2" value="Procurar">
                      </font> </td>
                  </tr> 
                </table>
        </form>
</td>
            </tr>


        <%
end if
'Verifica se a Session tem alguma posição, se tiver mostra a opção de apagar todos os vetores
'If ubound(vet) >= 0 Then
'Response.Write "<br>" &"<a href='vetor.asp?action=LimpaVetor'>Apagar Tudo</a>" & "<br>" 'Imprime o Vetor na tela
'End if

End Function

Function LimpaVetor2

'Limpa todas as posiçoes do vetor, apagando a Session
session("GuardaVetor") = Empty

End Function







'///////////////////////////////////////// vetor alunos /////////////////////////////////////////////////////////////////
Function VetorMontaAluno(Acao,Valor)
'Usamos o case para manipular a ação da função
Select Case Trim(Acao)
'Inclui nova posicao ao vetor
Case "Incluir"
'Guarda na variavel Vetor o conteudo da Session
Vetor = Session("GuardaVetor")
'Verifica se a Variavel Vetor é um Array, caso nao for entao definimos ela um Array
If Not IsArray(Vetor) Then Vetor = Array() End if
'Verifica se o Valor que esta sendo inserido já esta no Vetor se estiver entao nao inseri para nao haver duplicidades do vetor
If InStr(Join(Vetor), Valor) = 0 Then
'Este comando ira preservar o vetor e adciona + 1 valor
ReDim preserve Vetor(UBound(Vetor)+1)
'Este é o valor que estamos adicionando no vetor
Vetor(Ubound(Vetor )) = Valor
'Coloca o conteudo da variavel vetor dentro da Session
Session("GuardaVetor") = Vetor
End if
'Apaga uma determinada posicao do vetor
Case "Excluir"
'Inicia a varivel vetor como vazia
Vetor = Array()
'Criamos uma nova variavel Auxiliar e guardamos o valor da Session
AuxVetor = Session("GuardaVetor")
'Definine a Session como um Array vazio
Session("GuardaVetor") = Array()
'Faz um laço em todas as posições do vetor
For i = 0 To Ubound(AuxVetor)
'Verifica se o valor passado para excluir do vetor é diferente do valor que esta dentro da variavel Auxiliar
If AuxVetor(i) <> (Valor) Then
'Este comando ira preservar o vetor e adciona + 1 valor
ReDim preserve Vetor (UBound(Vetor)+1)
'Este é o valor que estamos adicionando no vetor
Vetor (Ubound(Vetor)) = AuxVetor(i)
'Coloca o conteudo da variavel vetor dentro da Session
Session("GuardaVetor") = Vetor
End If
Next
'Fim do Case
End Select
End Function

Function Incluir_Vetor_Aluno

'Executa a função que ira criar uma posição do vetor, basta passar a acao e o valor
Call VetorMontaAluno("Incluir",Valor_Vetor)
'Request("Valor_Vetor")
'response.Write(Valor_Vetor&"=vet<BR>")
End Function


Function VisualizaValoresVetorAluno
vet = session("GuardaVetor")

'Veriofica se a Session é um array, caso nao for então atribuimos a Session como um Array
IF Not IsArray(vet) Then vet = Array() End if
'Faremos um laço entre todos os vetores criados

if ubound(vet) >0 then
%>
  <tr> 
    <td valign="top"> 
<table width="1000" border="0" cellspacing="0">
  <tr> 
    <td valign="top"> 
      
      <table width="100%" border="0" align="right" cellspacing="0">
        <tr class="tb_corpo"
> 
          <td class="tb_tit"
>Lista de Alunos</td>
        </tr>
        <tr> 
          <td> <ul>
              <%
For x = 0 To ubound(vet) 

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos where CO_Matricula = "&vet(x)&" order BY NO_Aluno"
		RS.Open SQL, CON1

cod_cons =vet(x) 
no_aluno = RS("NO_Aluno")
Response.Write("<li><a class=ativos href =altera.asp?ori=1&cod_cons="&cod_cons&"&nvg="&nvg&">"&no_aluno&"</a></li>")
Next
%>
            </ul></td>
        </tr>
      </table></td>
  </tr>
</table>
<%

elseif ubound(vet)=0 then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos where NO_Aluno like '%"& strProcura & "%' order BY NO_Aluno"
		RS.Open SQL, CON1


cod_cons = RS("CO_Matricula")

response.Redirect("altera.asp?or=01&cod_cons="&cod_cons&"&nvg="&nvg)
else
Session("nome_cadastrar")=strProcura
%>
            <tr> 
              
    <td height="10" colspan="5"> 
      <%call mensagens(nivel,mensagem,1,0) %>
    </td>
			   </tr>
            <tr>           
    <td height="10" colspan="5"> 
      <%call mensagens(nivel,300,0,0) %>
    </td>
			  </tr>

        <tr> 
            <td valign="top"> 			  
        <form action="index.asp?opt=list&nvg=<%=nvg%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
          <table width="1000" border="0" cellspacing="0">
            <tr> 		
                  <tr class="tb_tit"> 			  
                    
      <td height="10" colspan="5">Preencha um dos campos abaixo</td>
                  </tr>
                  <tr> 
                    
      <td width="10"  height="10"> 
        <div align="right"><font class="form_dado_texto"> Matr&iacute;cula: 
          </font></div></td>
                    
      <td width="10"  height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
        </font><font size="2" face="Arial, Helvetica, sans-serif"> 
        <input name="busca1" type="text" class="textInput" id="busca1" size="12">
                      </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
                      </font></td>
                    
      <td width="10" height="10"> 
        <div align="right"><font class="form_dado_texto"> Nome: 
                        </font></div></td>
                    
      <td width="10"  height="10" ><font size="2" face="Arial, Helvetica, sans-serif"> 
        <input name="busca2" type="text" class="textInput" id="busca2" size="55" maxlength="50">
                      </font></td>
                    
      <td width="10" height="10" ><font size="2" face="Arial, Helvetica, sans-serif"> 
        <input name="Submit3" type="submit" class="botao_prosseguir" id="Submit2" value="Procurar">
                      </font> </td>
                  </tr> 
                </table>
        </form>
</td>
            </tr>

        <%
end if
'Verifica se a Session tem alguma posição, se tiver mostra a opção de apagar todos os vetores
'If ubound(vet) >= 0 Then
'Response.Write "<br>" &"<a href='vetor.asp?action=LimpaVetor'>Apagar Tudo</a>" & "<br>" 'Imprime o Vetor na tela
'End if

End Function

Function LimpaVetor2

'Limpa todas as posiçoes do vetor, apagando a Session
session("GuardaVetor") = Empty

End Function




'///////////////////////////////////////// vetor Web Família /////////////////////////////////////////////////////////////////
Function VetorMonta3(Acao,Valor)
'Usamos o case para manipular a ação da função
Select Case Trim(Acao)
'Inclui nova posicao ao vetor
Case "Incluir"
'Guarda na variavel Vetor o conteudo da Session
Vetor = Session("GuardaVetor")
'Verifica se a Variavel Vetor é um Array, caso nao for entao definimos ela um Array
If Not IsArray(Vetor) Then Vetor = Array() End if
'Verifica se o Valor que esta sendo inserido já esta no Vetor se estiver entao nao inseri para nao haver duplicidades do vetor
If InStr(Join(Vetor), Valor) = 0 Then
'Este comando ira preservar o vetor e adciona + 1 valor
ReDim preserve Vetor(UBound(Vetor)+1)
'Este é o valor que estamos adicionando no vetor
Vetor(Ubound(Vetor )) = Valor
'Coloca o conteudo da variavel vetor dentro da Session
Session("GuardaVetor") = Vetor
End if
'Apaga uma determinada posicao do vetor
Case "Excluir"
'Inicia a varivel vetor como vazia
Vetor = Array()
'Criamos uma nova variavel Auxiliar e guardamos o valor da Session
AuxVetor = Session("GuardaVetor")
'Definine a Session como um Array vazio
Session("GuardaVetor") = Array()
'Faz um laço em todas as posições do vetor
For i = 0 To Ubound(AuxVetor)
'Verifica se o valor passado para excluir do vetor é diferente do valor que esta dentro da variavel Auxiliar
If AuxVetor(i) <> (Valor) Then
'Este comando ira preservar o vetor e adciona + 1 valor
ReDim preserve Vetor (UBound(Vetor)+1)
'Este é o valor que estamos adicionando no vetor
Vetor (Ubound(Vetor)) = AuxVetor(i)
'Coloca o conteudo da variavel vetor dentro da Session
Session("GuardaVetor") = Vetor
End If
Next
'Fim do Case
End Select
End Function

Function Incluir_Vetor3

'Executa a função que ira criar uma posição do vetor, basta passar a acao e o valor
Call VetorMonta3("Incluir",Valor_Vetor)
'Request("Valor_Vetor")
'response.Write(Valor_Vetor&"=vet<BR>")
End Function


Function VisualizaValoresVetor3
vet = session("GuardaVetor")

'Veriofica se a Session é um array, caso nao for então atribuimos a Session como um Array
IF Not IsArray(vet) Then vet = Array() End if
'Faremos um laço entre todos os vetores criados

if ubound(vet) >0 then
%>
  <tr> 
    <td valign="top"> 
<table width="1000" border="0" cellspacing="0">
          <tr> 
            
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,70,0,0) %>
    </td>
          </tr>
  <tr> 
    <td valign="top"> 
      
      <table width="100%" border="0" align="right" cellspacing="0">
        <tr class="tb_corpo"
> 
          <td class="tb_tit"
>Lista de Usuários</td>
        </tr>
        <tr> 
          <td> <ul>
              <%
For x = 0 To ubound(vet) 

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Usuario where NO_Usuario = '"&vet(x)&"' order BY NO_Usuario"
		RS.Open SQL, CON_WF


cod_cons = RS("CO_Usuario")
Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=ativos href =altera.asp?ori=01&cod_cons="&cod_cons&"&nvg="&nvg&">"&vet(x)&"</a></font></li>")
Next
%>
            </ul></td>
        </tr>
      </table></td>
  </tr>
</table>
<%

elseif ubound(vet)=0 then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Usuario where NO_Usuario like '%"& strProcura & "%' order BY NO_Usuario"
		RS.Open SQL, CON_WF


cod_cons = RS("CO_Usuario")

response.Redirect("altera.asp?or=01&cod="&cod_cons&"&nvg="&nvg)
else
%>
            <tr> 
              
    <td height="10" colspan="5"> 
      <%call mensagens(nivel,mensagem,1,0) %>
    </td>
			   </tr>
            <tr> 
              
    <td height="10" colspan="5"> 
      <%call mensagens(nivel,69,0,0) %>
    </td>
			  </tr>
<form action="index.asp?opt=list&nvg=<%=nvg%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()"> 
          <tr class="tb_tit"> 
            
      <td height="10" colspan="5">Preencha um dos campos abaixo</td>
          </tr>
          <tr> 
            
      <td width="10"  height="10"> 
        <div align="right"><font class="form_dado_texto"> Usu&aacute;rio:</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
          </strong></font></div></td>
            
      <td width="10" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
        </font><font size="2" face="Arial, Helvetica, sans-serif"> 
        <input name="busca1" type="text" class="textInput" id="busca1" size="12">
        </font></td>
            
      <td width="10" height="10"> 
        <div align="right"><font class="form_dado_texto"> 
                Nome: </font></div></td>
            
      <td width="10" height="10" ><font size="2" face="Arial, Helvetica, sans-serif"> 
        <input name="busca2" type="text" class="textInput" id="busca2" size="55" maxlength="50">
              </font></td>
            
      <td width="10" height="10"><font size="2" face="Arial, Helvetica, sans-serif"> 
        <input name="Submit" type="submit" class="botao_prosseguir" id="Submit" value="Procurar">
              </font> </td>
          </tr>
</form>
 <tr>             
      <td > 
	  </td>
          </tr>
<%
end if
'Verifica se a Session tem alguma posição, se tiver mostra a opção de apagar todos os vetores
'If ubound(vet) >= 0 Then
'Response.Write "<br>" &"<a href='vetor.asp?action=LimpaVetor'>Apagar Tudo</a>" & "<br>" 'Imprime o Vetor na tela
'End if

End Function

Function LimpaVetor3

'Limpa todas as posiçoes do vetor, apagando a Session
session("GuardaVetor") = Empty

End Function


Function VisualizaValoresVetor4(tipo)
vet = session("GuardaVetor")
If tipo = "F" then
 titulo = "Fornecedores"
Elseif tipo = "I" then
 titulo = "Itens"
End if 

'Veriofica se a Session é um array, caso nao for então atribuimos a Session como um Array
IF Not IsArray(vet) Then vet = Array() End if
'Faremos um laço entre todos os vetores criados

if ubound(vet) >0 then
%>
<table width="1000" border="0" cellspacing="0">
  <tr> 
    <td valign="top"> 
      <table width="100%" border="0" align="right" cellspacing="0">
        <tr class="tb_corpo"
> 
          <td class="tb_tit"
>Lista de <%response.Write(titulo)%></td>
        </tr>
        <tr> 
          <td> <ul>
              <%
For x = 0 To ubound(vet) 
	if tipo = "F" then
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Fornecedor where NO_Fornecedor = '"&vet(x)&"' order BY NO_Fornecedor"
		RS.Open SQL, CON9


		cod_cons = RS("CO_Fornecedor")
		ativo = RS("IN_Ativo")
		if ativo = "True" then
			Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=ativos href=altera.asp?ch=ok&ori=01&cod_cons="&cod_cons&"&nvg="&nvg&" >"&vet(x)&"</a></font></li>")
		else
			Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=inativos href=altera.asp?ch=ok&ori=01&cod_cons="&cod_cons&"&nvg="&nvg&">"&vet(x)&"</a></font></li>")
		end if
		
	elseif tipo = "I" then		
	
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Item where NO_Item = '"&vet(x)&"' order BY NO_Item"
		RS.Open SQL, CON9

		cod_cons = RS("CO_Item")	
		
		Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=ativos href=altera.asp?ch=ok&ori=01&cod_cons="&cod_cons&"&nvg="&nvg&" >"&vet(x)&"</a></font></li>")				
	end if	
'Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a href =altera.asp?or=02&cod="&cod&">"&vet(x)&"</a></font></li>")
Next
%>
            </ul></td>
        </tr>
      </table></td>
  </tr>
</table>
<%

elseif ubound(vet)=0 then

	if tipo = "F" then
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Fornecedor where NO_Fornecedor = '"&vet(x)&"' order BY NO_Fornecedor"
		RS.Open SQL, CON9



		cod_cons = RS("CO_Fornecedor")		
	elseif tipo = "I" then		
	
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Item where NO_Item = '"&vet(x)&"' order BY NO_Item"
		RS.Open SQL, CON9



		cod_cons = RS("CO_Item")		
	end if	

	response.Redirect("altera.asp?ori=01&cod_cons="&cod_cons&"&nvg="&nvg)	
else

response.Redirect("index.asp?ori=01&opt=err2&cod_cons="&cod_cons&"&nvg="&nvg)%>

<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<%
end if
'Verifica se a Session tem alguma posição, se tiver mostra a opção de apagar todos os vetores
'If ubound(vet) >= 0 Then
'Response.Write "<br>" &"<a href='vetor.asp?action=LimpaVetor'>Apagar Tudo</a>" & "<br>" 'Imprime o Vetor na tela
'End if

End Function
%>