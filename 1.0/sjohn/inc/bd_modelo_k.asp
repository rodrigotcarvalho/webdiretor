<!--#include file="funcoes6.asp"-->
<%


	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CONg = Server.CreateObject("ADODB.Connection") 
	ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONg.Open ABRIRg		

function ReplicaInformacoes(unidade, curso, etapa, turma, matricula, periodo, modelo, campo, valor)

	Set RSRI = Server.CreateObject("ADODB.Recordset")
	SQLRI = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"'"
	RSRI.Open SQLRI, CON0
	
	tabela_nota = tabela_notas(CONg, unidade, curso, etapa, turma, null, null, null)
	caminho_nota = caminho_notas(CONg, tabela_nota, outro)
	
'	Set CONNT = Server.CreateObject("ADODB.Connection") 
'	ABRIRNT = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
'	CONNT.Open ABRIRNT	
	
	Dim CONNT,strSQL,objExec  
	Set CONNT = Server.Createobject("ADODB.Connection")  
	CONNT.Open "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}" 
  
On Error Resume Next '*** Error Resume ***' 	

CONNT.BeginTrans  
		
while not RSRI.EOF 	
	co_materia_principal = RSRI("CO_Materia")
'	response.Write(" P: "&co_materia_principal)

	wrk_tipo_materia = tipo_materia(co_materia_principal, curso, etapa)
	'response.Write(" T: "&wrk_tipo_materia)
	if wrk_tipo_materia = "T_F_F_N" or wrk_tipo_materia = "T_T_F_N" or wrk_tipo_materia = "T_F_T_N" then
		wrk_busca_filhas = busca_materias_filhas(co_materia_principal)
	'response.Write(" F: "&wrk_busca_filhas&"<BR>")		
		vetor_filhas = split(wrk_busca_filhas,"#!#")
		for vfls=0 to ubound(vetor_filhas)
			co_materia_filha = vetor_filhas(vfls)		
			Set RSNA = Server.CreateObject("ADODB.Recordset")
			SQLNA = "SELECT * FROM "&tabela_nota&" where CO_Matricula ="& matricula&" AND CO_Materia_Principal='"&co_materia_principal&"' AND CO_Materia='"&co_materia_filha&"' AND NU_Periodo = "&periodo		
			RSNA.Open SQLNA, CONNT
			
			if RSNA.EOF then
				IF ISNULL(valor) OR valor="" THEN
					valor = NULL
				end if
						
					Set RSNT = server.createobject("adodb.recordset")		
					RSNT.open tabela_nota, CONNT, 2, 2 'which table do you want open
					RSNT.addnew
				
					RSNT("CO_Matricula") = matricula
					RSNT("CO_Materia_Principal") = co_materia_principal
					RSNT("CO_Materia") = co_materia_filha
					RSNT("NU_Periodo") = periodo	
					RSNT(campo)=valor
						
					RSNT.update
					set RSNT=nothing	
						
			else
				IF ISNULL(valor) OR valor="" THEN
				  strSQL="UPDATE "&tabela_nota&" SET "&campo&" = NULL"
				ELSE
				  strSQL="UPDATE "&tabela_nota&" SET "&campo&" = "&valor
				END IF  		  

			  strSQL=strSQL & " WHERE CO_Matricula ="& matricula &" AND CO_Materia_Principal='"&co_materia_principal&"' AND CO_Materia='"&co_materia_filha&"' AND NU_Periodo = "&periodo	
 	
			  'CONNT.Execute sql
				Set objExec = CONNT.Execute(strSQL)  
				Set objExec = Nothing  			  
			end if
		Next	
	end if
RSRI.MOVENEXT
WEND
If Err.Number = 0 Then  
	'*** Commit Transaction ***'  
	CONNT.CommitTrans  
	Response.write("Save Done.")  
Else  
	'*** Rollback Transaction ***'  
	CONNT.RollbackTrans  
	Response.write("Error Save ["&strSQL&"] ("&Err.Description&")")  
End If  
  
CONNT.Close()  
Set objExec = Nothing  
Set CONNT = Nothing  
'	response.end()
end function

%>