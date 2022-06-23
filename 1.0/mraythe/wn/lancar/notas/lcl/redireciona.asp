<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/parametros.asp"-->
<!--#include file="../../../../inc/utils.asp"-->
<!--#include file="../../../../inc/bd_parametros.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->

<%
vetorDatas = request.form("data_form")
submit = request.form("submit")


	if submit="Alterar" then
		vetor_data_alterar = split(vetorDatas,", ")
		response.Redirect("alterar.asp?acao=a&P_DATA_AULA="&vetor_data_alterar(0))

	else
		session("obr") = vetorDatas
		response.Redirect("confirmar.asp")	
	end if

%>