<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->





<!--#include file="../../../../inc/caminhos.asp"-->


<% 
Response.buffer = True
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
cod= request.QueryString("cod_cons")
opt = request.QueryString("opt")
	

obr=cod


enturma = request.QueryString("et")
recria_at = SESSION("recria_at")
recria_att = SESSION("recria_att")

'response.Write(">>"&enturma)

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")		

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON1_aux = Server.CreateObject("ADODB.Connection") 
		ABRIR1_aux = "DBQ="& CAMINHO_al_aux & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1_aux.Open ABRIR1_aux
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2				
					
	
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		
		
if enturma="att" then		
	Set RS1 = Server.CreateObject("ADODB.Recordset")
	if recria_att="s" then		
		SQL1 = "SELECT * FROM TB_Matriculas WHERE NU_Ano="&ano_letivo&" order by CO_Matricula"
	else
		SQL1 = "SELECT * FROM TB_Matriculas WHERE CO_Situacao='P' AND NU_Ano="&ano_letivo&" order by CO_Matricula"
	end if				
	RS1.Open SQL1, CON1
'RESPONSE.Write(SQL1&"<br>")
		IF RS1.EOF THEN		
		ELSE%>
          <%
			WHile Not RS1.EOF
			unidade_aluno = RS1("NU_Unidade")
			curso_aluno = RS1("CO_Curso")
			etapa_aluno = RS1("CO_Etapa")
			turma_aluno = RS1("CO_Turma")
			
			Set RS1_aux = Server.CreateObject("ADODB.Recordset")
			SQL1_aux = "SELECT * FROM TBI_Enturmar_Aluno WHERE NU_Unidade="&unidade_aluno&" AND CO_Curso='"&curso_aluno&"' AND CO_Etapa='"&etapa_aluno&"' AND CO_Turma='"&turma_aluno&"'"
			RS1_aux.Open SQL1_aux, CON1_aux
'RESPONSE.Write("<br><br>"&SQL1_aux)
			
			if RS1_aux.eof then
							
				Set RS1_toda = Server.CreateObject("ADODB.Recordset")
				SQL1_toda = "SELECT * FROM TB_Matriculas WHERE NU_Ano="&ano_letivo&" AND NU_Unidade="&unidade_aluno&" AND CO_Curso='"&curso_aluno&"' AND CO_Etapa='"&etapa_aluno&"' AND CO_Turma='"&turma_aluno&"' Order BY CO_Matricula"
				RS1_toda.Open SQL1_toda, CON1
'RESPONSE.Write("<br>"&SQL1_toda)				
				WHile Not RS1_toda.EOF					
									
					cod = RS1_toda("CO_Matricula")
					chamada = RS1_toda("NU_Chamada")
					matricula = RS1_toda("DA_Rematricula")
					unidade_pesquisa = RS1_toda("NU_Unidade")
					curso_pesquisa = RS1_toda("CO_Curso")
					etapa_pesquisa = RS1_toda("CO_Etapa")
					turma_pesquisa = RS1_toda("CO_Turma")
					situacao = RS1_toda("CO_Situacao")
								
						Set RS2 = Server.CreateObject("ADODB.Recordset")
						SQL2 = "SELECT * FROM TB_Alunos WHERE CO_Matricula="&cod
						RS2.Open SQL2, CON1
						
						nome = RS2("NO_Aluno")								
											
					Set RSALUNO_bd = server.createobject("adodb.recordset")
					RSALUNO_bd.open "TBI_Enturmar_Aluno", CON1_aux, 2, 2
					RSALUNO_bd.addnew				
					RSALUNO_bd("CO_Matricula")=cod
					RSALUNO_bd("NO_Aluno")=nome
					RSALUNO_bd("NU_Chamada")=chamada
					RSALUNO_bd("CO_Situacao")=situacao								
					RSALUNO_bd("DA_Rematricula")=data_matricula				
					RSALUNO_bd("NU_Unidade")=unidade_pesquisa				
					RSALUNO_bd("CO_Curso")=curso_pesquisa		
					RSALUNO_bd("CO_Etapa")=etapa_pesquisa
					RSALUNO_bd("CO_Turma")=turma_pesquisa					
					RSALUNO_bd.update
				RS1_toda.Movenext
				wend								
					if recria_att="s" then
						Set RS1_ENT = Server.CreateObject("ADODB.Recordset")
						SQL1_ENT = "SELECT * FROM TBI_Enturmar_Aluno WHERE NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='"&etapa_pesquisa&"' AND CO_Turma='"&turma_pesquisa&"' Order BY NO_Aluno"
						RS1_ENT.Open SQL1_ENT, CON1_aux								
					
						nu_chamada=0
						WHile Not RS1_ENT.EOF
							cod = RS1_ENT("CO_Matricula")
							nome = RS1_ENT("NO_Aluno")
							nu_chamada = nu_chamada+1
							situacao = RS1_ENT("CO_Situacao")				
							
							if situacao="P" then
								sql_mat="UPDATE TB_Matriculas SET CO_Situacao='C', NU_Chamada='"&nu_chamada&"' where NU_Ano="&ano_letivo&" and CO_Matricula="&cod
								Set RS_mat = CON1.Execute(sql_mat)
								set RS_mat=nothing
							else
								sql_mat="UPDATE TB_Matriculas SET NU_Chamada='"&nu_chamada&"' where NU_Ano="&ano_letivo&" and CO_Matricula="&cod
								Set RS_mat = CON1.Execute(sql_mat)
								set RS_mat=nothing
							end if
							
							sql_AxT="UPDATE TB_Aluno_Esta_Turma SET NU_Chamada='"&nu_chamada&"' where CO_Matricula="&cod
							Set RS_AxT = CON2.Execute(sql_AxT)
							set RS_AxT=nothing
						'Response.Write("<BR>U "&unidade_aluno&" - "&curso_aluno&" - "&etapa_aluno&" - "&turma_aluno&" / ch> "&nu_chamada&" cod> "&cod&" nome> "&nome&" - "&situacao&"</a>")
						  					  
						RS1_ENT.Movenext
						Wend
					'	Response.Flush						
					else
					
						Set RS1_chamada = Server.CreateObject("ADODB.Recordset")
						SQL1_chamada = "SELECT MAX(NU_Chamada) AS chamada TBI_Enturmar_Aluno WHERE NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='"&etapa_pesquisa&"' AND CO_Turma='"&turma_pesquisa&"'"
						RS1_chamada.Open SQL1_chamada, CON1
					
							IF RS1_chamada.EOF then
							nu_chamada =0
							else 	
							nu_chamada = RS1_chamada("chamada")		
							end if					
																	
						Set RS1_ENT = Server.CreateObject("ADODB.Recordset")
						SQL1_ENT = "SELECT * FROM TBI_Enturmar_Aluno WHERE CO_Situacao='P' AND NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='"&etapa_pesquisa&"' AND CO_Turma='"&turma_pesquisa&"' Order BY DA_Rematricula"
						RS1_ENT.Open SQL1_ENT, CON1_aux								
					
						WHile Not RS1_ENT.EOF
							cod = RS1_ENT("CO_Matricula")
							nome = RS1_ENT("NO_Aluno")
							nu_chamada = nu_chamada+1			
							
							sql_mat="UPDATE TB_Matriculas SET CO_Situacao='C', NU_Chamada='"&nu_chamada&"' where NU_Ano="&ano_letivo&" and CO_Matricula="&cod
							Set RS_mat = CON1.Execute(sql_mat)
							set RS_mat=nothing
							
							sql_AxT="UPDATE TB_Aluno_Esta_Turma SET NU_Chamada='"&nu_chamada&"' where CO_Matricula="&cod
							Set RS_AxT = CON2.Execute(sql_AxT)
							set RS_AxT=nothing
						  					  
						RS1_ENT.Movenext
						Wend
					'	Response.Flush
					END IF

			RS1.Movenext				
			ELSE
			RS1.Movenext
			END IF
			wend
			Set RSCONTATO = Server.CreateObject("ADODB.Recordset")
			SQLAA= "DELETE * FROM TBI_Enturmar_Aluno"
			RSCONTATO.Open SQLAA, CON1_aux			
		END IF	
'Response.Clear()		
response.Redirect("index.asp?nvg=WS-MA-MA-ETA&opt=ok")

elseif enturma="at" then
unidade_pesquisa=SESSION("unidade_pesquisa")
curso_pesquisa=SESSION("curso_pesquisa")
etapa_pesquisa=SESSION("etapa_pesquisa")
turma_pesquisa=SESSION("turma_pesquisa")

'				SQL1 = "SELECT * FROM TB_Matriculas WHERE NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='"&etapa_pesquisa&"' AND CO_Turma='"&turma_pesquisa&"' AND NU_Ano="&ano_letivo
'response.Write(SQL1)
'response.End()

if unidade_pesquisa="999990" then
sql_1=""
sql_2=""
sql_3=""
else
sql_1="NU_Unidade="&unidade_pesquisa&" AND"
sql_2="AND NU_Unidade="&unidade_pesquisa
sql_3=" WHERE NU_Unidade="&unidade_pesquisa
end if

if curso_pesquisa="999990" then
sql_1=""
sql_2=sql_2
sql_3=sql_3
else
sql_1="NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND"
sql_2="AND NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"'"
sql_3=" WHERE NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"'"
end if

if etapa_pesquisa="999990" then
sql_1=""
sql_2=sql_2
sql_3=sql_3
else
sql_1="NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='"&etapa_pesquisa&"' AND"
sql_2="AND NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='"&etapa_pesquisa&"'"
sql_3=" WHERE NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='"&etapa_pesquisa&"'"
end if

if turma_pesquisa="999990" then
sql_1=""
sql_2=sql_2
sql_3=sql_3
else
sql_1="NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='"&etapa_pesquisa&"' AND CO_Turma='"&turma_pesquisa&"' AND"
sql_2= "AND NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='"&etapa_pesquisa&"' AND CO_Turma='"&turma_pesquisa&"'"
sql_3= " WHERE NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='"&etapa_pesquisa&"' AND CO_Turma='"&turma_pesquisa&"'"
end if
	Set RS1 = Server.CreateObject("ADODB.Recordset")
	if recria_at="s" then		
		SQL1 = "SELECT * FROM TB_Matriculas WHERE "&sql_1&" NU_Ano="&ano_letivo&" order by CO_Matricula"
	else
		SQL1 = "SELECT * FROM TB_Matriculas WHERE CO_Situacao='P' "&sql_2&" AND NU_Ano="&ano_letivo&" order by CO_Matricula"
	end if				
	RS1.Open SQL1, CON1
'RESPONSE.Write(SQL1&"<br>")
		IF RS1.EOF THEN		
		ELSE%>
          <%
			WHile Not RS1.EOF
			unidade_aluno = RS1("NU_Unidade")
			curso_aluno = RS1("CO_Curso")
			etapa_aluno = RS1("CO_Etapa")
			turma_aluno = RS1("CO_Turma")
			
			Set RS1_aux = Server.CreateObject("ADODB.Recordset")
			SQL1_aux = "SELECT * FROM TBI_Enturmar_Aluno WHERE NU_Unidade="&unidade_aluno&" AND CO_Curso='"&curso_aluno&"' AND CO_Etapa='"&etapa_aluno&"' AND CO_Turma='"&turma_aluno&"'"
			RS1_aux.Open SQL1_aux, CON1_aux
'RESPONSE.Write("<br><br>"&SQL1_aux)
			
			if RS1_aux.eof then
							
				Set RS1_toda = Server.CreateObject("ADODB.Recordset")
				SQL1_toda = "SELECT * FROM TB_Matriculas WHERE NU_Ano="&ano_letivo&" AND NU_Unidade="&unidade_aluno&" AND CO_Curso='"&curso_aluno&"' AND CO_Etapa='"&etapa_aluno&"' AND CO_Turma='"&turma_aluno&"' Order BY CO_Matricula"
				RS1_toda.Open SQL1_toda, CON1
'RESPONSE.Write("<br>"&SQL1_toda)				
				WHile Not RS1_toda.EOF					
									
					cod = RS1_toda("CO_Matricula")
					chamada = RS1_toda("NU_Chamada")
					matricula = RS1_toda("DA_Rematricula")
					unidade_pesquisa = RS1_toda("NU_Unidade")
					curso_pesquisa = RS1_toda("CO_Curso")
					etapa_pesquisa = RS1_toda("CO_Etapa")
					turma_pesquisa = RS1_toda("CO_Turma")
					situacao = RS1_toda("CO_Situacao")
								
						Set RS2 = Server.CreateObject("ADODB.Recordset")
						SQL2 = "SELECT * FROM TB_Alunos WHERE CO_Matricula="&cod
						RS2.Open SQL2, CON1
						
						nome = RS2("NO_Aluno")								
											
					Set RSALUNO_bd = server.createobject("adodb.recordset")
					RSALUNO_bd.open "TBI_Enturmar_Aluno", CON1_aux, 2, 2
					RSALUNO_bd.addnew				
					RSALUNO_bd("CO_Matricula")=cod
					RSALUNO_bd("NO_Aluno")=nome
					RSALUNO_bd("NU_Chamada")=chamada
					RSALUNO_bd("CO_Situacao")=situacao								
					RSALUNO_bd("DA_Rematricula")=data_matricula				
					RSALUNO_bd("NU_Unidade")=unidade_pesquisa				
					RSALUNO_bd("CO_Curso")=curso_pesquisa		
					RSALUNO_bd("CO_Etapa")=etapa_pesquisa
					RSALUNO_bd("CO_Turma")=turma_pesquisa					
					RSALUNO_bd.update
				RS1_toda.Movenext
				wend								
					if recria_at="s" then
						Set RS1_ENT = Server.CreateObject("ADODB.Recordset")
						SQL1_ENT = "SELECT * FROM TBI_Enturmar_Aluno WHERE NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='"&etapa_pesquisa&"' AND CO_Turma='"&turma_pesquisa&"' Order BY NO_Aluno"
						RS1_ENT.Open SQL1_ENT, CON1_aux								
					
						nu_chamada=0
						WHile Not RS1_ENT.EOF
							cod = RS1_ENT("CO_Matricula")
							nome = RS1_ENT("NO_Aluno")
							nu_chamada = nu_chamada+1
							situacao = RS1_ENT("CO_Situacao")				
							
							if situacao="P" then
								sql_mat="UPDATE TB_Matriculas SET CO_Situacao='C', NU_Chamada='"&nu_chamada&"' where NU_Ano="&ano_letivo&" and CO_Matricula="&cod
								Set RS_mat = CON1.Execute(sql_mat)
								set RS_mat=nothing
							else
								sql_mat="UPDATE TB_Matriculas SET NU_Chamada='"&nu_chamada&"' where NU_Ano="&ano_letivo&" and CO_Matricula="&cod
								Set RS_mat = CON1.Execute(sql_mat)
								set RS_mat=nothing
							end if
							
							sql_AxT="UPDATE TB_Aluno_Esta_Turma SET NU_Chamada='"&nu_chamada&"' where CO_Matricula="&cod
							Set RS_AxT = CON2.Execute(sql_AxT)
							set RS_AxT=nothing
						'Response.Write("<BR>U "&unidade_aluno&" - "&curso_aluno&" - "&etapa_aluno&" - "&turma_aluno&" / ch> "&nu_chamada&" cod> "&cod&" nome> "&nome&" - "&situacao&"</a>")
						  					  
						RS1_ENT.Movenext
						Wend
					'	Response.Flush						
					else
					
						Set RS1_chamada = Server.CreateObject("ADODB.Recordset")
						SQL1_chamada = "SELECT MAX(NU_Chamada) AS chamada TBI_Enturmar_Aluno WHERE NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='"&etapa_pesquisa&"' AND CO_Turma='"&turma_pesquisa&"'"
						RS1_chamada.Open SQL1_chamada, CON1
					
							IF RS1_chamada.EOF then
							nu_chamada =0
							else 	
							nu_chamada = RS1_chamada("chamada")		
							end if					
																	
						Set RS1_ENT = Server.CreateObject("ADODB.Recordset")
						SQL1_ENT = "SELECT * FROM TBI_Enturmar_Aluno WHERE CO_Situacao='P' AND NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='"&etapa_pesquisa&"' AND CO_Turma='"&turma_pesquisa&"' Order BY DA_Rematricula"
						RS1_ENT.Open SQL1_ENT, CON1_aux								
					
						WHile Not RS1_ENT.EOF
							cod = RS1_ENT("CO_Matricula")
							nome = RS1_ENT("NO_Aluno")
							nu_chamada = nu_chamada+1			
							
							sql_mat="UPDATE TB_Matriculas SET CO_Situacao='C', NU_Chamada='"&nu_chamada&"' where NU_Ano="&ano_letivo&" and CO_Matricula="&cod
							Set RS_mat = CON1.Execute(sql_mat)
							set RS_mat=nothing
							
							sql_AxT="UPDATE TB_Aluno_Esta_Turma SET NU_Chamada='"&nu_chamada&"' where CO_Matricula="&cod
							Set RS_AxT = CON2.Execute(sql_AxT)
							set RS_AxT=nothing
						  					  
						RS1_ENT.Movenext
						Wend
						'Response.Flush
					END IF

			RS1.Movenext				
			ELSE
			RS1.Movenext
			END IF
			wend
			Set RSCONTATO = Server.CreateObject("ADODB.Recordset")
			SQLAA= "DELETE * FROM TBI_Enturmar_Aluno"
			RSCONTATO.Open SQLAA, CON1_aux			
		END IF			
'response.end()		
response.Redirect("index.asp?nvg=WS-MA-MA-ETA&opt=ok2")




elseif enturma="mt" then
unidade_pesquisa=SESSION("unidade_pesquisa")
curso_pesquisa=SESSION("curso_pesquisa")
etapa_pesquisa=SESSION("etapa_pesquisa")
turma_pesquisa=SESSION("turma_pesquisa")

Session("GuardaMatriculas") = Array()
Session("GuardaChamadasDigitadas") = Array()
Session("GuardaSituacoes") = Array()
Session("ValidaChamadas") = Array()

'armazena valores digitados
	Set RS1 = Server.CreateObject("ADODB.Recordset")
	SQL1 = "SELECT * FROM TB_Matriculas WHERE NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='"&etapa_pesquisa&"' AND CO_Turma='"&turma_pesquisa&"' AND NU_Ano="&ano_letivo&" order by NU_Chamada"
	RS1.Open SQL1, CON1
	
	
	
	WHile Not RS1.EOF				
	matricula = RS1("CO_Matricula")
	chamada=request.Form("chamada_"&matricula)
	situacao=request.Form("situac_"&matricula)				
	
		matriculas = Session("GuardaMatriculas")
		If Not IsArray(matriculas) Then matriculas = Array() End if
		If InStr(Join(matriculas), matricula) = 0 Then
		ReDim preserve matriculas(UBound(matriculas)+1)
		matriculas(Ubound(matriculas)) = matricula
		Session("GuardaMatriculas") = matriculas
		End if
		
		chamadas = Session("GuardaChamadasDigitadas")
		If Not IsArray(chamadas) Then chamadas = Array() End if
		ReDim preserve chamadas(UBound(chamadas)+1)
		chamadas(Ubound(chamadas)) = chamada
		Session("GuardaChamadasDigitadas") = chamadas
		
		situacoes = Session("GuardaSituacoes")
		If Not IsArray(situacoes) Then situacoes = Array() End if
		ReDim preserve situacoes(UBound(situacoes)+1)
		situacoes(Ubound(situacoes)) = situacao
		Session("GuardaSituacoes") = situacoes
		

	
	RS1.Movenext
	Wend

'Verifica se existem números de chamadas duplicadas
	Set RS1 = Server.CreateObject("ADODB.Recordset")
	SQL1 = "SELECT * FROM TB_Matriculas WHERE NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='"&etapa_pesquisa&"' AND CO_Turma='"&turma_pesquisa&"' AND NU_Ano="&ano_letivo&" order by NU_Chamada"
	RS1.Open SQL1, CON1

	valida_chamadas = Session("ValidaChamadas")
	'for i=0 to ubound(valida_chamadas)					
	'response.write("<br>vc_"&ubound(valida_chamadas))
	'next
'i=-1
	WHile Not RS1.EOF
	'i=i+1				
	matricula = RS1("CO_Matricula")
	chamada=request.Form("chamada_"&matricula)
	situacao=request.Form("situac_"&matricula)


	'response.write("<br>_"&chamada)


		If Not IsArray(valida_chamadas) Then valida_chamadas = Array() End if
'é necessário colocar aspas simples no número de chamda pois o sistema confunde 30 com 3 e diz que o número já existe no vetor
		If InStr(Join(valida_chamadas), "'"&chamada&"'") = 0 Then
		ReDim preserve valida_chamadas(UBound(valida_chamadas)+1)
		valida_chamadas(Ubound(valida_chamadas)) = "'"&chamada&"'"
		Session("ValidaChamadas") = valida_chamadas
		else
		response.Redirect("altera.asp?opt=err1")
		End if
'		response.write("<br>+"&valida_chamadas(i))
'		response.Write(" vc "&UBound(valida_chamadas))
	RS1.Movenext
	Wend				
'response.End()

	for i=0 to	Ubound(matriculas)
	
	cod=matriculas(i)
	nu_chamada_vetor=valida_chamadas(i)
'é necessário esse split pois o número de chamada vem no vetor entre aspas simples ex:'1','2' etc	
	nu_chamada_split=split(nu_chamada_vetor,"'")
	nu_chamada=nu_chamada_split(1)
	situacao=situacoes(i)
					'response.Write(cod&" - "&nu_chamada&" - "&situacao&"<br>")
							
					if situacao="P" then
						sql_mat="UPDATE TB_Matriculas SET CO_Situacao='C', NU_Chamada='"&nu_chamada&"' where NU_Ano="&ano_letivo&" and CO_Matricula="&cod
						Set RS_mat = CON1.Execute(sql_mat)
						set RS_mat=nothing
					else
						sql_mat="UPDATE TB_Matriculas SET NU_Chamada='"&nu_chamada&"' where NU_Ano="&ano_letivo&" and CO_Matricula="&cod
						Set RS_mat = CON1.Execute(sql_mat)
						set RS_mat=nothing
					end if
					
					sql_AxT="UPDATE TB_Aluno_Esta_Turma SET NU_Chamada='"&nu_chamada&"' where CO_Matricula="&cod
					Set RS_AxT = CON2.Execute(sql_AxT)
					set RS_AxT=nothing
	next
response.Redirect("index.asp?nvg=WS-MA-MA-ETA&opt=ok3")												
end if

If Err.number<>0 then
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