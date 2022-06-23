<%'On Error Resume Next%>



<!--#include file="../../../../inc/caminhos.asp"-->
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

	if id="n" then
	nome="cidnat"
	elseif id="r" then
	nome="cidres"
		if familiar="n" then
		javascript="onChange=recuperarBairroRes('"&uf&"',this.value)"
		else
	nome="cidres_fam"	
		javascript="onChange=recuperarBairroResFam('"&uf&"',this.value)"
		end if
	elseif id="c" then
	nome="cidcom"
		if familiar="n" then
		javascript="onChange=recuperarBairroCom('"&uf&"',this.value)"
		else
	nome="cidcom_fam"	
		javascript="onChange=recuperarBairroComFam('"&uf&"',this.value)"
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
end if

elseif id="c" then
if familiar="n" then
nome="bairrocom"
else
nome="bairrocom_fam"
end if
end if

uf= request.form("c_pub")
cidade= request.form("b_pub")

'response.Write(">>"&uf&"<<"&cidade)
%>
  <select name="<%response.Write(nome)%>" class="textInput" id="select">
<%
	



		Set RS6b = Server.CreateObject("ADODB.Recordset")
		SQL6b = "SELECT * FROM TB_Bairros WHERE SG_UF ='"& uf&"' AND CO_Municipio="& cidade&" order by NO_Bairro"
		RS6b.Open SQL6b, CON0
		
IF RS6b.EOF then
%>		
                      <option value="100"> 
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