<%'On Error Resume Next%>
<!--#include file="../../../../inc/connect_al.asp"-->
<!--#include file="../../../../inc/connect_pr.asp"-->
<!--#include file="../../../../inc/connect_ct.asp"-->
<!--#include file="../../../../inc/connect_l.asp"-->
<% 
opt= request.querystring("opt")
id= request.querystring("o")
familiar= request.querystring("f")

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR		

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0


		
if opt="c" then
uf= request.form("c_pub")
cod= request.form("cod_pub")
cod_familiar=request.form("c_fam_pub")
tp_vinc_familiar_aux=request.form("tp_vinc_pub")
co_vinc_familiar_aux=request.form("c_vinc_pub")

	if id="n" then
	nome="cidnat"
	elseif id="r" then
	nome="cidres"
	session("uf_res")=uf
	session("uf_com")=session("uf_com")
		if familiar="n" then
		javascript="onChange='recuperarBairroRes(this.value)'"
		else
	nome="cidres_fam"	
		javascript="onChange=recuperarBairroResFam(this.value);BD_aux(this.value,cod_consulta.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'CO_Municipio_Res')"
		end if
	elseif id="c" then
	nome="cidcom"
	session("uf_com")=uf
	session("uf_res")=session("uf_res")
		if familiar="n" then
		javascript="onChange='recuperarBairroCom(this.value)'"
		else
	nome="cidcom_fam"	
		javascript="onChange=recuperarBairroComFam(this.value);BD_aux(this.value,cod_consulta.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'CO_Municipio_Com')"
		end if
	elseif id="e" then
	nome="cid_curs"
	javascript=""
end if

%>
  <select name="<%response.Write(nome)%>" class="textInput" id="<%response.Write(nome)%>" <%response.Write(javascript)%>>
  				                        <option value="0"></option>
<%
	


		Set RS6 = Server.CreateObject("ADODB.Recordset")
		SQL6 = "SELECT * FROM TB_Municipios WHERE SG_UF ='"& uf&"' order by NO_Municipio"
		RS6.Open SQL6, CON0

while not RS6.EOF						
natural= RS6("CO_Municipio")
NO_UF= RS6("NO_Municipio")
%>
                      <option value="<%response.Write(Server.URLEncode(natural))%>"> 
                      <% response.Write(Server.URLEncode(NO_UF))%>
                      </option>
                      <%
RS6.MOVENEXT
WEND
%>
                    </select>
<% elseif opt="b" then
if id="r" then
if familiar="n" then
nome="bairrores"
else
nome="bairrores_fam"
javascript="onChange=BD_aux(this.value,cod_consulta.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'CO_Bairro_Res')"
end if
uf_sessao=session("uf_res")
elseif id="c" then
if familiar="n" then
nome="bairrocom"
else
nome="bairrocom_fam"
javascript="onChange=BD_aux(this.value,cod_consulta.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'CO_Bairro_Com')"
end if
uf_sessao=session("uf_com")
end if

%>
  <select name="<%response.Write(nome)%>" class="textInput" id="select" <%response.Write(javascript)%>>
<%
cidade= request.form("b_pub")	



		Set RS6b = Server.CreateObject("ADODB.Recordset")
		SQL6b = "SELECT * FROM TB_Bairros WHERE SG_UF ='"& uf_sessao&"' AND CO_Municipio="& cidade&" order by NO_Bairro"
		RS6b.Open SQL6b, CON0
		
IF RS6b.EOF then
%>		
                      <option value="0"> 
                      <% response.Write(Server.URLEncode("Bairros não cadastrados"))%>
                      </option>
<%else					  
while not RS6b.EOF						
co_bairro= RS6b("CO_Bairro")
nome_bairro= RS6b("NO_Bairro")
%>
                      <option value="<%response.Write(Server.URLEncode(co_bairro))%>"> 
                      <% response.Write(Server.URLEncode(nome_bairro))%>
                      </option>
                      <%
RS6b.MOVENEXT
WEND
end if
%>
                    </select>


<%end if%>
<%If Err.number<>0 then
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