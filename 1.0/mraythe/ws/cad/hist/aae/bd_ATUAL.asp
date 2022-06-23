<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<%
chave=session("nvg")
session("nvg")=chave

opt = request.QueryString("opt")


Set CON7 = Server.CreateObject("ADODB.Connection") 
ABRIR7 = "DBQ="& CAMINHO_h & ";Driver={Microsoft Access Driver (*.mdb)}"
CON7.Open ABRIR7	


ano_letivo = session("ano_letivo")
co_usr = session("co_user")

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

if mes<10 then
meswrt="0"&mes
else
meswrt=mes
end if
if min<10 then
minwrt="0"&min
else
minwrt=min
end if

data_log = dia &"/"& meswrt &"/"& ano&" "&hora & ":"& minwrt
data_bd = dia &"/"& meswrt &"/"& ano
if opt="exc" then
	exclui_historico=request.form("exclui_historico")				
	vetorExclui = split(exclui_historico,",")
	for i =0 to ubound(vetorExclui)

		exclui = split(vetorExclui(i),"$!$")
		
			ano_historico = exclui(0)
			nu_seq_hist = exclui(1)
			cod_aluno  = exclui(2)	
	
			Set RS = Server.CreateObject("ADODB.Recordset")
			SQL = "DELETE * from TB_Historico_Ano where CO_Matricula = "& cod_aluno &" AND DA_Ano = "& ano_historico&" AND NU_Seq = "& nu_seq_hist
			RS.Open SQL, CON7
			
			
			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL2 = "DELETE * from TB_Historico_Nota where CO_Matricula = "& cod_aluno &" AND DA_Ano = "& ano_historico&" AND NU_Seq = "& nu_seq_hist
			RS2.Open SQL2, CON7			
	
	next
	obr=session("obr")
	session("obr")=obr
	
	
	outro= "Excluir,"&data_log&","&replace(exclui_historico,"$!$","-")
	call GravaLog (chave,outro)
	response.redirect("resumo.asp?voltar=S&opt=ok3")

elseif opt="inc" or opt="alt" then

	historico= request.form("dados_historico")
	ano_hist_form= request.form("ano_hist_form")	
	tipo_curso=request.form("tipo_curso")	
	co_seg=request.form("co_seg")	
	estabelecimento_form=request.form("estabelecimento_form")	
	pais_form=request.form("pais_form")
	uf_form=request.form("uf_form")
	municipio_form= request.form("municipio_form")
	resultado_final_form= request.form("resultado_final_form")
	observacoes_form= request.form("observacoes_form")
	carga_total_form= request.form("carga_total_form")
	dias_letivos_form= request.form("dias_letivos_form")	
	frequencia_total_form= request.form("frequencia_total_form")
	qtd_itens= request.form("qtd_itens")		
	dados_historico = 	split(historico,"$!$")
	
	if ubound(dados_historico)<2 then
		ano_hist = ano_hist_form
		seq_hist = 1
		matric_hist = dados_historico(0)
	else
		ano_hist = dados_historico(0)
		seq_hist = dados_historico(1)
		matric_hist	= dados_historico(2)	
	end if
	

	
	if opt="inc" then	
		outro= "Incluir "&historico
	elseif opt="alt" then
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "DELETE * from TB_Historico_Ano where CO_Matricula = "& matric_hist &" AND DA_Ano = "& ano_hist&" AND NU_Seq = "& seq_hist
		RS.Open SQL, CON7
		
		
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "DELETE * from TB_Historico_Nota where CO_Matricula = "& matric_hist &" AND DA_Ano = "& ano_hist&" AND NU_Seq = "& seq_hist
		RS2.Open SQL2, CON7		
		
		outro= "Alterar "&historico
	end if

	
		
	if carga_total_form="" then
		carga_total_form = NULL
	end if
	
	

	if frequencia_total_form="" then
		frequencia_total_form = NULL
	end if	
	

	
	if qtd_itens="" then
		qtd_itens = NULL
	end if	
	
	Set RS = server.createobject("adodb.recordset")
	RS.open "TB_Historico_Ano", CON7, 2, 2 'which table do you want open
	RS.addnew	
		RS("CO_Matricula") = matric_hist
		RS("DA_Ano") = ano_hist_form
		RS("NU_Seq") = seq_hist
		RS("TP_Curso") = tipo_curso
		RS("CO_Seg") = co_seg
		RS("NO_Escola") = estabelecimento_form
		RS("NO_Pais") = pais_form
		RS("NO_Municipio") = municipio_form
		RS("SG_UF") = uf_form
		RS("IN_Aprovado") = resultado_final_form
		RS("TX_Observacoes") = observacoes_form
		RS("NU_ANO_Letivo") = NULL
		RS("TP_Registro") = "M"
		RS("DT_Registro") = data_bd
		RS("NU_Carga_Horaria_Total") = carga_total_form
		RS("NU_Dias_Letivo") = dias_letivos_form
		RS("TX_Frequencia_Total") = frequencia_total_form
	RS.update	  
	set RS=nothing
	
	
	
	for frm = 1 to qtd_itens
		disciplina= request.form("disciplina_"&frm)
		carga_form_disc= request.form("carga_form_"&frm)
		frequencia_form_disc= request.form("frequencia_form_"&frm)
		nota_form_disc= request.form("nota_form_"&frm)
		aprovado_disc= request.form("aprovado_"&frm)
		
		if aprovado_disc = "S" then
			aprovado_disc = TRUE
		else
			aprovado_disc = FALSE
		end if
		
		if carga_form_disc="" then
			carga_form_disc = NULL
		end if			
		
		if frequencia_form_disc="" then
			frequencia_form_disc = NULL
		end if	
			response.Write(qtd_itens&"-"&frm&"-"&matric_hist&"-"&ano_hist_form&"-"&seq_hist&"-"&disciplina&"<BR>")			
		if (not isnull(disciplina)) and disciplina<>"" then	
			response.Write("OK<BR>")
			Set RS = server.createobject("adodb.recordset")
			RS.open "TB_Historico_Nota", CON7, 2, 2 'which table do you want open
			RS.addnew	
			
				RS("CO_Matricula") = matric_hist
				RS("DA_Ano") = ano_hist_form
				RS("NU_Seq") = seq_hist
				RS("NO_Materia") = disciplina
				RS("NU_Carga_Horaria") = carga_form_disc
				RS("TX_Frequencia") = frequencia_form_disc
				RS("VA_Nota") = nota_form_disc
				RS("IN_Aprovado") = aprovado_disc
			RS.update	  
			set RS=nothing		
		end if
	Next
	
	call GravaLog (chave,outro)	
	
	response.redirect("incluir.asp?opt="&opt&"&res=ok&cod="&historico)


end if



%>