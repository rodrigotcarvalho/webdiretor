<!--#include file="caminhos.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="funcoes2.asp"-->
<!--#include file="funcoes6.asp"-->
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
Function mapa_notas (CAMINHOa,CAMINHOn,unidade,curso,etapa,turma,co_materia,periodo,ano_letivo,co_usr,opcao,erro,print)
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
 mes = "mar�o"
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
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20" class="tb_subtit">&nbsp;</td>
    <td width="300" class="tb_subtit"> 
      <div align="left"></div></td>
    <td colspan="12" class="tb_tit"> <div align="center">Aproveitamento</div></td>
    <td colspan="4" class="tb_tit"><div align="center">Freq&uuml;&ecirc;ncia (Faltas)</div></td>
  </tr>
  <tr>
    <td width="20" class="tb_subtit">N&ordm;</td>
    <td width="300" class="tb_subtit">Nome</td>
    <td width="68" class="tb_subtit"> <div align="center">BIM 1</div></td>
    <td width="68" class="tb_subtit"> <div align="center">BIM 2</div></td>
    <td width="68" class="tb_subtit"><div align="center">M&eacute;dia Sem 1</div></td>
    <td width="68" class="tb_subtit"><div align="center">Recup Sem</div></td>
    <td width="68" class="tb_subtit"><div align="center">M&eacute;dia Sem 2</div></td>
    <td width="68" class="tb_subtit"> <div align="center">BIM 3</div></td>
    <td width="68" class="tb_subtit"> <div align="center">BIM 4</div></td>
    <td width="68" class="tb_subtit"><div align="center">M&eacute;dia Sem 3</div></td>
    <td width="68" class="tb_subtit"> <div align="center">M&eacute;dia Anual</div></td>
    <td width="68" class="tb_subtit"> <div align="center">Recup Final </div></td>
    <td width="68" class="tb_subtit"><div align="center">M&eacute;dia Final </div></td>
    <td width="68" class="tb_subtit"><div align="center">Result</div></td>
    <td width="68" class="tb_subtit"><div align="center">BIM 1</div></td>
    <td width="68" class="tb_subtit"> <div align="center">BIM 2</div></td>
    <td width="68" class="tb_subtit"> <div align="center">BIM 3</div></td>
    <td width="68" class="tb_subtit"> <div align="center">BIM 4</div></td>
  </tr>
  <%
			rec_lancado="sim"
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
				
		Set CON_A = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_A.Open ABRIR
				
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)
		
			check=2
		
			select case opcao
				case 1
				tb="TB_NOTA_A"
				case 2
				tb="TB_NOTA_B"
				case 3
				tb="TB_NOTA_C"
			end select
					
		while not RS.EOF
			
			nu_matricula = RS("CO_Matricula")
			session("matricula")=nu_matricula
			nu_chamada = RS("NU_Chamada")
  
  		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
		Set RS2 = CON_A.Execute(SQL_A)
	
		no_aluno= RS2("NO_Aluno")
			
					dividendo1=0
					divisor1=0
					dividendo2=0
					divisor2=0
					dividendo3=0
					divisor3=0
					dividendo4=0
					divisor4=0
					dividendo_ma=0
					divisor_ma=0
					dividendo5=0
					divisor5=0
					dividendo_mf=0
					divisor_mf=0
					dividendo6=0
					divisor6=0
					dividendo_rec=0
					divisor_rec=0
					res1="&nbsp;"
					res2="&nbsp;"
					res3="&nbsp;"
					m2="&nbsp;"
					m3="&nbsp;"										
					media_final="&nbsp;"
					resultado_final="&nbsp;"			
			
				Set RS1a = Server.CreateObject("ADODB.Recordset")
				SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia='"&co_materia&"'"
				RS1a.Open SQL1a, CON0
					
				no_materia=RS1a("NO_Materia")
			
				if check mod 2 =0 then
				cor = "tb_fundo_linha_par" 
				cor2 = "tb_fundo_linha_impar" 				
				else 
				cor ="tb_fundo_linha_impar"
				cor2 = "tb_fundo_linha_par" 
				end if
			
					
					
				Set CON_N = Server.CreateObject("ADODB.Connection") 
				ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
				CON_N.Open ABRIRn
			
						for periodofil=1 to 5
											
								
							Set RSnFIL = Server.CreateObject("ADODB.Recordset")
							Set RS3 = Server.CreateObject("ADODB.Recordset")
							SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodofil
							Set RS3 = CON_N.Execute(SQL_N)
						
						
						
							if RS3.EOF then
								if periodofil=1 then
								f1="&nbsp;"
								va_m31="&nbsp;"
								elseif periodofil=2 then
								f2="&nbsp;"
								va_m32="&nbsp;"
								va_rec_sem="&nbsp;"
								elseif periodofil=3 then
								f3="&nbsp;"
								va_m33="&nbsp;"
								elseif periodofil=4 then
								f4="&nbsp;"
								va_m34="&nbsp;"
								elseif periodofil=5 then
								va_m35="&nbsp;"
								elseif periodofil=6 then
								va_m36="&nbsp;"
								end if	
							else
								if periodofil=1 then
								f1=RS3("NU_Faltas")
								va_m31=RS3("VA_Media3")
								elseif periodofil=2 then
								f2=RS3("NU_Faltas")
								va_m32=RS3("VA_Media3")
								va_rec_sem=RS3("VA_Rec")
								elseif periodofil=3 then
								f3=RS3("NU_Faltas")
								va_m33=RS3("VA_Media3")
								elseif periodofil=4 then
								f4=RS3("NU_Faltas")
								va_m34=RS3("VA_Media3")
								elseif periodofil=5 then
								va_m35=RS3("VA_Media3")
								elseif periodofil=6 then
								va_m36=RS3("VA_Media3")
								end if
							end if
						NEXT
							
							if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
							dividendo1=0
							divisor1=0
							va_m31="&nbsp;"
							else
							dividendo1=va_m31
							divisor1=1
							end if	
							
							if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
							dividendo2=0
							divisor2=0
							va_m32="&nbsp;"
							else
							dividendo2=va_m32
							divisor2=1
							end if
							
							if isnull(va_rec_sem) or va_rec_sem="&nbsp;"  or va_rec_sem="" then
							dividendorec=0
							divisorrec=0
							va_rec_sem="&nbsp;" 
							else
							dividendorec=va_rec_sem
							divisorrec=1
							end if
							
							if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
							dividendo3=0
							divisor3=0
							va_m33="&nbsp;"
							else
							dividendo3=va_m33
							divisor3=1
							end if
							
							if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
							dividendo4=0
							divisor4=0
							va_m34="&nbsp;"
							else
							dividendo4=va_m34
							divisor4=1
							end if														
							if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
							nota_aux_m2_1="&nbsp;"
							'dividendo5=0
							'divisor5=0
							va_m35="&nbsp;"
							else
							nota_aux_m2_1=va_m35
							'dividendo5=va_m35
							'divisor5=1
							end if

									showapr1="s"
									showprova1="s"
									showapr2="s"
									showprova2="s"
									showapr3="s"
									showprova3="s"
									showapr4="s"
									showprova4="s"

							dividendo_ms1=dividendo1+dividendo2
							divisor_ms1=divisor1+divisor2
													
							if divisor_ms1<2 then
							ms1="&nbsp;"
							dividendoms1=0
							divisorms1=0
							else
							ms1=dividendo_ms1/divisor_ms1
								decimo = ms1 - Int(ms1)
								If decimo >= 0.5 Then
									nota_arredondada = Int(ms1) + 1
									ms1=nota_arredondada
								'elseIf decimo >= 0.25 Then
								'	nota_arredondada = Int(ms1) + 0.5
								'	ms1=nota_arredondada
								else
									nota_arredondada = Int(ms1)
									ms1=nota_arredondada											
								End If			
							ms1 = formatNumber(ms1,0)
							dividendoms1=ms1
							divisorms1=1
							end if
							

							
							if divisorrec=0 then
								ms2=ms1
								if ms2="&nbsp;" then
									dividendoms2=0
									divisorms2=0
									dividendo_anual_ms2=0
									divisor_anual_ms2=0
								else
									dividendoms2=ms2
									divisorms2=1						
									dividendo_anual_ms2=ms2
									divisor_anual_ms2=1
								end if
							else
								dividendo_ms2=dividendoms1+dividendorec
								divisor_ms2=divisorms1+divisorrec
																						
								ms2=dividendo_ms2/divisor_ms2
'response.Write(ms2&"+"&dividendoms1&"+"&divisor_ms2)
ms2=ms2*1	
ms1=ms1*1							
								if ms2<ms1 then
									ms2=ms1								
								end if
									decimo = ms2 - Int(ms2)
									If decimo >= 0.5 Then
										nota_arredondada = Int(ms2) + 1
										ms2=nota_arredondada
									'elseIf decimo >= 0.25 Then
									'	nota_arredondada = Int(ms2) + 0.5
									'	ms2=nota_arredondada
									else
										nota_arredondada = Int(ms2)
										ms2=nota_arredondada											
									End If
								ms2 = formatNumber(ms2,0)																
								dividendo_anual_ms2=ms2
								divisor_anual_ms2=1
							end if
							
							dividendo_ms3=dividendo3+dividendo4
							divisor_ms3=divisor3+divisor4
							
							if divisor_ms3<2 then
							ms3="&nbsp;"
							dividendo_anual_ms3=0
							divisor_anual_ms3=0					
							else
							ms3=dividendo_ms3/divisor_ms3
								decimo = ms3 - Int(ms3)
								If decimo >= 0.5 Then
									nota_arredondada = Int(ms3) + 1
									ms3=nota_arredondada
								'elseIf decimo >= 0.25 Then
								'	nota_arredondada = Int(ms3) + 0.5
								'	ms3=nota_arredondada
								else
									nota_arredondada = Int(ms3)
									ms3=nota_arredondada											
								End If
								ms3 = formatNumber(ms3,0)								
							dividendo_anual_ms3=ms3
							divisor_anual_ms3=1						
							end if					
								dividendo_anual_ms2=dividendo_anual_ms2*1
								dividendo_anual_ms3=dividendo_anual_ms3*1
								divisor_anual_ms2=divisor_anual_ms2*1
								divisor_anual_ms3=divisor_anual_ms3*1		
							dividendo_ma=dividendo_anual_ms2+dividendo_anual_ms3
							divisor_ma=divisor_anual_ms2+divisor_anual_ms3
							
							'response.Write(dividendo_ma&"<<")
			'response.Write(ma&"<<BR>")						
							if divisor_ma<2 then
							ma="&nbsp;"
							else
							ma=dividendo_ma/divisor_ma
							end if
							
							if ma="&nbsp;" then
							else
							'mf=mf/10
								decimo = ma - Int(ma)
		
									If decimo >= 0.5 Then
										nota_arredondada = Int(ma) + 1
										ma=nota_arredondada
									'elseIf decimo >= 0.25 Then
									'	nota_arredondada = Int(ma) + 0.5
									'	ma=nota_arredondada								
									Else
										nota_arredondada = Int(ma)
										ma=nota_arredondada					
									End If
									
								ma = formatNumber(ma,0)
							end if
	
							if ma="&nbsp;" then
								media_final="&nbsp;"
								resultado_final="&nbsp;"							
							else	
							
								if nota_aux_m2_1="&nbsp;" then
									tipo_calculo="anual"
								else
									tipo_calculo="recuperacao"
								end if	
								
								if ma<>"" and not isnull(ma) then	
									ma_regra = ma/10
								end if									
	
								resultado=regra_aprovacao(curso,etapa,ma_regra,nota_aux_m2_1,"&nbsp;","&nbsp;","&nbsp;",tipo_calculo)
								resultado_aluno = split(resultado,"#!#")
								media_final=resultado_aluno(0)
								resultado_final=resultado_aluno(1)								
							end if
							
							
'response.Write(Session("resultado_1") &" - "& Session("resultado_2") &" - "& Session("resultado_3")&"<<")
			%>
  <tr> 
    <td width="20" class="<%response.Write(cor)%>"> 
      <%response.Write(nu_chamada)
						  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
						  %>
    </td>
    <td width="300" class="<%response.Write(cor)%>"> 
      <%response.Write(no_aluno)
						  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
						  %>
    </td>
    <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
        <%
							if showapr1="s" then																	
							response.Write(va_m31)
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
        <%
							if showapr1="s" then												
							response.Write(va_m32)						
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
        <%
							if showapr3="s" then					
							response.Write(ms1)
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
        <%
							if showapr1="s" then					
							response.Write(va_rec_sem)
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
        <%
							if showapr3="s" then					
							response.Write(ms2)
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
        <%
							if showapr1="s" then					
							response.Write(va_m33)
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
        <%
							if showapr1="s" then					
							response.Write(va_m34)
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
        <%
							if showapr3="s" then					
							response.Write(ms3)
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
        <%
							if showapr3="s" then
							response.Write(ma)
							else
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
        <%
							if showapr2="s" and showprova2="s" then												
							response.Write(va_m35)
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"><div align="center"><%				
								response.Write(media_final)
							%></div></td>
    <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
                          <%				
								response.Write(resultado_final)
							%>
                        </div></td>
    <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
        <%
							response.Write("&nbsp;"&f1)
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
        <%
							response.Write("&nbsp;"&f2)
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
        <%
							response.Write("&nbsp;"&f3)
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
        <%
							response.Write("&nbsp;"&f4)
							%>
      </div></td>
  </tr>
  <%
					check=check+1
			RS.MOVENEXT
			wend		
			%>
</table>

<%
End Function














'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function imp_mapa_notas (CAMINHOa,CAMINHOn,unidade,curso,etapa,turma,co_materia,periodo,ano_letivo,co_usr,opcao,erro,print)

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
 mes = "mar�o"
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
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20" class="tabelaTit">&nbsp;</td>
    <td width="300" class="tabelaTit"> 
      <div align="left">&nbsp;</div></td>
    <td colspan="12" class="tabelaTit"> <div align="center">Aproveitamento</div></td>
    <td colspan="4" class="tabelaTit"><div align="center">Freq&uuml;&ecirc;ncia (Faltas)</div></td>
  </tr>
  <tr>
    <td width="20" class="tabelaTit">N&ordm;</td>
    <td width="300" class="tabelaTit">Nome</td>
    <td width="68" class="tabelaTit"> <div align="center">BIM 1</div></td>
    <td width="68" class="tabelaTit"> <div align="center">BIM 2</div></td>
    <td width="68" class="tabelaTit"><div align="center">M&eacute;dia Sem 1</div></td>
    <td width="68" class="tabelaTit"><div align="center">Recup Sem</div></td>
    <td width="68" class="tabelaTit"><div align="center">M&eacute;dia Sem 2</div></td>
    <td width="68" class="tabelaTit"> <div align="center">BIM 3</div></td>
    <td width="68" class="tabelaTit"> <div align="center">BIM 4</div></td>
    <td width="68" class="tabelaTit"><div align="center">M&eacute;dia Sem 3</div></td>
    <td width="68" class="tabelaTit"> <div align="center">M&eacute;dia Anual</div></td>
    <td width="68" class="tabelaTit"> <div align="center">Recup Final </div></td>
    <td width="68" class="tabelaTit"> <div align="center">M&eacute;dia Final</div></td>
    <td width="68" class="tabelaTit"><div align="center">Result</div></td>
    <td width="68" class="tabelaTit"><div align="center">BIM 1</div></td>
    <td width="68" class="tabelaTit"> <div align="center">BIM 2</div></td>
    <td width="68" class="tabelaTit"> <div align="center">BIM 3</div></td>
    <td width="68" class="tabelaTit"> <div align="center">BIM 4</div></td>
  </tr>
  <%
			rec_lancado="sim"
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
				
		Set CON_A = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_A.Open ABRIR
				
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)
		
			check=2
		
			select case opcao
				case 1
				tb="TB_NOTA_A"
				case 2
				tb="TB_NOTA_B"
				case 3
				tb="TB_NOTA_C"
			end select
					
		while not RS.EOF
			
			nu_matricula = RS("CO_Matricula")
			session("matricula")=nu_matricula
			nu_chamada = RS("NU_Chamada")
  
  		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
		Set RS2 = CON_A.Execute(SQL_A)
	
		no_aluno= RS2("NO_Aluno")
			
					dividendo1=0
					divisor1=0
					dividendo2=0
					divisor2=0
					dividendo3=0
					divisor3=0
					dividendo4=0
					divisor4=0
					dividendo_ma=0
					divisor_ma=0
					dividendo5=0
					divisor5=0
					dividendo_mf=0
					divisor_mf=0
					dividendo6=0
					divisor6=0
					dividendo_rec=0
					divisor_rec=0
					res1="&nbsp;"
					res2="&nbsp;"
					res3="&nbsp;"
					m2="&nbsp;"
					m3="&nbsp;"										
					media_final="&nbsp;"
					resultado_final="&nbsp;"				
			
				Set RS1a = Server.CreateObject("ADODB.Recordset")
				SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia='"&co_materia&"'"
				RS1a.Open SQL1a, CON0
					
				no_materia=RS1a("NO_Materia")
			
				if check mod 2 =0 then
				cor = "tabela" 
				cor2 = "tabela" 				
				else 
				cor ="tabela"
				cor2 = "tabela" 
				end if
			
					
					
				Set CON_N = Server.CreateObject("ADODB.Connection") 
				ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
				CON_N.Open ABRIRn
			
						for periodofil=1 to 5
											
								
							Set RSnFIL = Server.CreateObject("ADODB.Recordset")
							Set RS3 = Server.CreateObject("ADODB.Recordset")
							SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodofil
							Set RS3 = CON_N.Execute(SQL_N)
						
						
						
							if RS3.EOF then
								if periodofil=1 then
								f1="&nbsp;"
								va_m31="&nbsp;"
								elseif periodofil=2 then
								f2="&nbsp;"
								va_m32="&nbsp;"
								va_rec_sem="&nbsp;"
								elseif periodofil=3 then
								f3="&nbsp;"
								va_m33="&nbsp;"
								elseif periodofil=4 then
								f4="&nbsp;"
								va_m34="&nbsp;"
								elseif periodofil=5 then
								va_m35="&nbsp;"
								elseif periodofil=6 then
								va_m36="&nbsp;"
								end if	
							else
								if periodofil=1 then
								f1=RS3("NU_Faltas")
								va_m31=RS3("VA_Media3")
								elseif periodofil=2 then
								f2=RS3("NU_Faltas")
								va_m32=RS3("VA_Media3")
								va_rec_sem=RS3("VA_Rec")
								elseif periodofil=3 then
								f3=RS3("NU_Faltas")
								va_m33=RS3("VA_Media3")
								elseif periodofil=4 then
								f4=RS3("NU_Faltas")
								va_m34=RS3("VA_Media3")
								elseif periodofil=5 then
								va_m35=RS3("VA_Media3")
								elseif periodofil=6 then
								va_m36=RS3("VA_Media3")
								end if
							end if
						NEXT
							
							if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
							dividendo1=0
							divisor1=0
							va_m31="&nbsp;"
							else
							dividendo1=va_m31
							divisor1=1
							end if	
							
							if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
							dividendo2=0
							divisor2=0
							va_m32="&nbsp;"
							else
							dividendo2=va_m32
							divisor2=1
							end if
							
							if isnull(va_rec_sem) or va_rec_sem="&nbsp;"  or va_rec_sem="" then
							dividendorec=0
							divisorrec=0
							va_rec_sem="&nbsp;" 
							else
							dividendorec=va_rec_sem
							divisorrec=1
							end if
							
							if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
							dividendo3=0
							divisor3=0
							va_m33="&nbsp;"
							else
							dividendo3=va_m33
							divisor3=1
							end if
							
							if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
							dividendo4=0
							divisor4=0
							va_m34="&nbsp;"
							else
							dividendo4=va_m34
							divisor4=1
							end if														
							if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
							nota_aux_m2_1="&nbsp;"
							'dividendo5=0
							'divisor5=0
							va_m35="&nbsp;"
							else
							nota_aux_m2_1=va_m35
							'dividendo5=va_m35
							'divisor5=1
							end if


									showapr1="s"
									showprova1="s"
									showapr2="s"
									showprova2="s"
									showapr3="s"
									showprova3="s"
									showapr4="s"
									showprova4="s"

							dividendo_ms1=dividendo1+dividendo2
							divisor_ms1=divisor1+divisor2
													
							if divisor_ms1<2 then
							ms1="&nbsp;"
							dividendoms1=0
							divisorms1=0
							else
							ms1=dividendo_ms1/divisor_ms1
								decimo = ms1 - Int(ms1)
								If decimo >= 0.5 Then
									nota_arredondada = Int(ms1) + 1
									ms1=nota_arredondada
								'elseIf decimo >= 0.25 Then
								'	nota_arredondada = Int(ms1) + 0.5
								'	ms1=nota_arredondada
								else
									nota_arredondada = Int(ms1)
									ms1=nota_arredondada											
								End If			
							ms1 = formatNumber(ms1,0)
							dividendoms1=ms1
							divisorms1=1
							end if
							

							
							if divisorrec=0 then
								ms2=ms1
								if ms2="&nbsp;" then
									dividendoms2=0
									divisorms2=0
									dividendo_anual_ms2=0
									divisor_anual_ms2=0
								else
									dividendoms2=ms2
									divisorms2=1						
									dividendo_anual_ms2=ms2
									divisor_anual_ms2=1
								end if
							else
								dividendo_ms2=dividendoms1+dividendorec
								divisor_ms2=divisorms1+divisorrec
																						
								ms2=dividendo_ms2/divisor_ms2
'response.Write(ms2&"+"&dividendoms1&"+"&divisor_ms2)
ms2=ms2*1	
ms1=ms1*1							
								if ms2<ms1 then
									ms2=ms1								
								end if
									decimo = ms2 - Int(ms2)
									If decimo >= 0.5 Then
										nota_arredondada = Int(ms2) + 1
										ms2=nota_arredondada
									'elseIf decimo >= 0.25 Then
									'	nota_arredondada = Int(ms2) + 0.5
									'	ms2=nota_arredondada
									else
										nota_arredondada = Int(ms2)
										ms2=nota_arredondada											
									End If
								ms2 = formatNumber(ms2,0)																
								dividendo_anual_ms2=ms2
								divisor_anual_ms2=1
							end if
							
							dividendo_ms3=dividendo3+dividendo4
							divisor_ms3=divisor3+divisor4
							
							if divisor_ms3<2 then
							ms3="&nbsp;"
							dividendo_anual_ms3=0
							divisor_anual_ms3=0					
							else
							ms3=dividendo_ms3/divisor_ms3
								decimo = ms3 - Int(ms3)
								If decimo >= 0.5 Then
									nota_arredondada = Int(ms3) + 1
									ms3=nota_arredondada
								'elseIf decimo >= 0.25 Then
								'	nota_arredondada = Int(ms3) + 0.5
								'	ms3=nota_arredondada
								else
									nota_arredondada = Int(ms3)
									ms3=nota_arredondada											
								End If
								ms3 = formatNumber(ms3,0)								
							dividendo_anual_ms3=ms3
							divisor_anual_ms3=1						
							end if					
								dividendo_anual_ms2=dividendo_anual_ms2*1
								dividendo_anual_ms3=dividendo_anual_ms3*1
								divisor_anual_ms2=divisor_anual_ms2*1
								divisor_anual_ms3=divisor_anual_ms3*1		
							dividendo_ma=dividendo_anual_ms2+dividendo_anual_ms3
							divisor_ma=divisor_anual_ms2+divisor_anual_ms3
							
							'response.Write(dividendo_ma&"<<")
			'response.Write(ma&"<<BR>")						
							if divisor_ma<2 then
							ma="&nbsp;"
							else
							ma=dividendo_ma/divisor_ma
							end if
							
							if ma="&nbsp;" then
							else
							'mf=mf/10
								decimo = ma - Int(ma)
		
									If decimo >= 0.5 Then
										nota_arredondada = Int(ma) + 1
										ma=nota_arredondada
									'elseIf decimo >= 0.25 Then
									'	nota_arredondada = Int(ma) + 0.5
									'	ma=nota_arredondada								
									Else
										nota_arredondada = Int(ma)
										ma=nota_arredondada					
									End If
									
								ma = formatNumber(ma,0)
							end if
	
							if ma="&nbsp;" then
								media_final="&nbsp;"
								resultado_final="&nbsp;"							
							else	
							
								if nota_aux_m2_1="&nbsp;" then
									tipo_calculo="anual"
								else
									tipo_calculo="recuperacao"
								end if	
								if ma<>"" and not isnull(ma) then	
									ma_regra = ma/10
								end if
								resultado=regra_aprovacao(curso,etapa,ma_regra,nota_aux_m2_1,"&nbsp;","&nbsp;","&nbsp;",tipo_calculo)
								resultado_aluno = split(resultado,"#!#")
								media_final=resultado_aluno(0)
								resultado_final=resultado_aluno(1)								
							end if
							
							
'response.Write(Session("resultado_1") &" - "& Session("resultado_2") &" - "& Session("resultado_3")&"<<")
			%>
  <tr> 
    <td width="20" class="<%response.Write(cor)%>"> 
      <%response.Write(nu_chamada)
						  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
						  %>
    </td>
    <td width="300" class="<%response.Write(cor)%>"> 
      <%response.Write(no_aluno)
						  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
						  %>
    </td>
    <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
        <%
							if showapr1="s" then																	
							response.Write(va_m31)
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
        <%
							if showapr1="s" then												
							response.Write(va_m32)						
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
        <%
							if showapr3="s" then					
							response.Write(ms1)
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
        <%
							if showapr1="s" then					
							response.Write(va_rec_sem)
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
        <%
							if showapr3="s" then					
							response.Write(ms2)
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
        <%
							if showapr1="s" then					
							response.Write(va_m33)
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
        <%
							if showapr1="s" then					
							response.Write(va_m34)
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
        <%
							if showapr3="s" then					
							response.Write(ms3)
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
        <%
							if showapr3="s" then
							response.Write(ma)
							else
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
        <%
							if showapr2="s" and showprova2="s" then												
							response.Write(va_m35)
							end if
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor)%>"><div align="center"><%				
								response.Write(media_final)
							%></div></td>
    <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                          <%				
								response.Write(resultado_final)
							%>
                        </div></td>
    <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
        <%
							response.Write("&nbsp;"&f1)
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
        <%
							response.Write("&nbsp;"&f2)
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
        <%
							response.Write("&nbsp;"&f3)
							%>
      </div></td>
    <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
        <%
							response.Write("&nbsp;"&f4)
							%>
      </div></td>
  </tr>
  <%
					check=check+1
			RS.MOVENEXT
			wend		
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
 mes = "mar�o"
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
<% response.Write("impresso em "& data &" �s "& horario)
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
<% response.Write("impresso em "& data &" �s "& horario)
%></font>
 </div>
</td></tr>
              </table>
<%
end if 
end function			  
%>

