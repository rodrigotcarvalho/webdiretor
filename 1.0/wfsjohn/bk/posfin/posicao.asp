<% 
Set CON0 = Server.CreateObject("ADODB.Connection") 
CAMINHO0 = "e:\home\bretanha\dados\logins.mdb"
ABRIR0 = "DBQ="& CAMINHO0 & ";Driver={Microsoft Access Driver (*.mdb)}"
CON0.Open ABRIR0
Set RS0 = Server.CreateObject("ADODB.Recordset")
CONEXAO0 = "SELECT * FROM logins WHERE login='" & Session("login") & "'"
RS0.Open CONEXAO0, CON0
co_mat = RS0("CO_Matricula_Escola")

user = Session("usuario")
url = "posicao2.asp?opt=" & co_mat

if (user ="A") Then
response.redirect(url)
else
response.redirect("posicaor.asp")
End IF
%>