<% 
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
cod_cons= request.QueryString("cod_cons")	
response.Redirect("mapa.asp?cod_cons="&cod_cons)
%>

