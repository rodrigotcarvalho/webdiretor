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
'///////////////////////////////////////////////    VETOR     //////////////////////////////////////////////////////////////////////////////


Function VetorMonta(Acao,Valor)
'Usamos o case para manipular a a��o da fun��o
Select Case Trim(Acao)
'Inclui nova posicao ao vetor
Case "Incluir"
'Guarda na variavel Vetor o conteudo da Session
Vetor = Session("GuardaVetor")
'Verifica se a Variavel Vetor � um Array, caso nao for entao definimos ela um Array
If Not IsArray(Vetor) Then Vetor = Array() End if
'Verifica se o Valor que esta sendo inserido j� esta no Vetor se estiver entao nao inseri para nao haver duplicidades do vetor
If InStr(Join(Vetor), Valor) = 0 Then
'Este comando ira preservar o vetor e adciona + 1 valor
ReDim preserve Vetor(UBound(Vetor)+1)
'Este � o valor que estamos adicionando no vetor
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
'Faz um la�o em todas as posi��es do vetor
For i = 0 To Ubound(AuxVetor)
'Verifica se o valor passado para excluir do vetor � diferente do valor que esta dentro da variavel Auxiliar
If AuxVetor(i) <> (Valor) Then
'Este comando ira preservar o vetor e adciona + 1 valor
ReDim preserve Vetor (UBound(Vetor)+1)
'Este � o valor que estamos adicionando no vetor
Vetor (Ubound(Vetor)) = AuxVetor(i)
'Coloca o conteudo da variavel vetor dentro da Session
Session("GuardaVetor") = Vetor
End If
Next
'Fim do Case
End Select
End Function

Function Incluir_Vetor

'Executa a fun��o que ira criar uma posi��o do vetor, basta passar a acao e o valor
Call VetorMonta("Incluir",Valor_Vetor)
'Request("Valor_Vetor")
'response.Write(Valor_Vetor&"=vet<BR>")
End Function


Function VisualizaValoresVetor
vet = session("GuardaVetor")

'Veriofica se a Session � um array, caso nao for ent�o atribuimos a Session como um Array
IF Not IsArray(vet) Then vet = Array() End if
'Faremos um la�o entre todos os vetores criados

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
'Verifica se a Session tem alguma posi��o, se tiver mostra a op��o de apagar todos os vetores
'If ubound(vet) >= 0 Then
'Response.Write "<br>" &"<a href='vetor.asp?action=LimpaVetor'>Apagar Tudo</a>" & "<br>" 'Imprime o Vetor na tela
'End if

End Function

Function LimpaVetor

'Limpa todas as posi�oes do vetor, apagando a Session
session("GuardaVetor") = Empty

End Function
'///////////////////////////////////////// vetor alunos /////////////////////////////////////////////////////////////////
Function VetorMonta2(Acao,Valor)
'Usamos o case para manipular a a��o da fun��o
Select Case Trim(Acao)
'Inclui nova posicao ao vetor
Case "Incluir"
'Guarda na variavel Vetor o conteudo da Session
Vetor = Session("GuardaVetor")
'Verifica se a Variavel Vetor � um Array, caso nao for entao definimos ela um Array
If Not IsArray(Vetor) Then Vetor = Array() End if
'Verifica se o Valor que esta sendo inserido j� esta no Vetor se estiver entao nao inseri para nao haver duplicidades do vetor
If InStr(Join(Vetor), Valor) = 0 Then
'Este comando ira preservar o vetor e adciona + 1 valor
ReDim preserve Vetor(UBound(Vetor)+1)
'Este � o valor que estamos adicionando no vetor
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
'Faz um la�o em todas as posi��es do vetor
For i = 0 To Ubound(AuxVetor)
'Verifica se o valor passado para excluir do vetor � diferente do valor que esta dentro da variavel Auxiliar
If AuxVetor(i) <> (Valor) Then
'Este comando ira preservar o vetor e adciona + 1 valor
ReDim preserve Vetor (UBound(Vetor)+1)
'Este � o valor que estamos adicionando no vetor
Vetor (Ubound(Vetor)) = AuxVetor(i)
'Coloca o conteudo da variavel vetor dentro da Session
Session("GuardaVetor") = Vetor
End If
Next
'Fim do Case
End Select
End Function

Function Incluir_Vetor2

'Executa a fun��o que ira criar uma posi��o do vetor, basta passar a acao e o valor
Call VetorMonta("Incluir",Valor_Vetor)
'Request("Valor_Vetor")
'response.Write(Valor_Vetor&"=vet<BR>")
End Function


Function VisualizaValoresVetor2
vet = session("GuardaVetor")

'Veriofica se a Session � um array, caso nao for ent�o atribuimos a Session como um Array
IF Not IsArray(vet) Then vet = Array() End if
'Faremos um la�o entre todos os vetores criados

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
'Verifica se a Session tem alguma posi��o, se tiver mostra a op��o de apagar todos os vetores
'If ubound(vet) >= 0 Then
'Response.Write "<br>" &"<a href='vetor.asp?action=LimpaVetor'>Apagar Tudo</a>" & "<br>" 'Imprime o Vetor na tela
'End if

End Function

Function LimpaVetor2

'Limpa todas as posi�oes do vetor, apagando a Session
session("GuardaVetor") = Empty

End Function







'///////////////////////////////////////// vetor alunos /////////////////////////////////////////////////////////////////
Function VetorMontaAluno(Acao,Valor)
'Usamos o case para manipular a a��o da fun��o
Select Case Trim(Acao)
'Inclui nova posicao ao vetor
Case "Incluir"
'Guarda na variavel Vetor o conteudo da Session
Vetor = Session("GuardaVetor")
'Verifica se a Variavel Vetor � um Array, caso nao for entao definimos ela um Array
If Not IsArray(Vetor) Then Vetor = Array() End if
'Verifica se o Valor que esta sendo inserido j� esta no Vetor se estiver entao nao inseri para nao haver duplicidades do vetor
If InStr(Join(Vetor), Valor) = 0 Then
'Este comando ira preservar o vetor e adciona + 1 valor
ReDim preserve Vetor(UBound(Vetor)+1)
'Este � o valor que estamos adicionando no vetor
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
'Faz um la�o em todas as posi��es do vetor
For i = 0 To Ubound(AuxVetor)
'Verifica se o valor passado para excluir do vetor � diferente do valor que esta dentro da variavel Auxiliar
If AuxVetor(i) <> (Valor) Then
'Este comando ira preservar o vetor e adciona + 1 valor
ReDim preserve Vetor (UBound(Vetor)+1)
'Este � o valor que estamos adicionando no vetor
Vetor (Ubound(Vetor)) = AuxVetor(i)
'Coloca o conteudo da variavel vetor dentro da Session
Session("GuardaVetor") = Vetor
End If
Next
'Fim do Case
End Select
End Function

Function Incluir_Vetor_Aluno

'Executa a fun��o que ira criar uma posi��o do vetor, basta passar a acao e o valor
Call VetorMontaAluno("Incluir",Valor_Vetor)
'Request("Valor_Vetor")
'response.Write(Valor_Vetor&"=vet<BR>")
End Function


Function VisualizaValoresVetorAluno
vet = session("GuardaVetor")

'Veriofica se a Session � um array, caso nao for ent�o atribuimos a Session como um Array
IF Not IsArray(vet) Then vet = Array() End if
'Faremos um la�o entre todos os vetores criados

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
'Verifica se a Session tem alguma posi��o, se tiver mostra a op��o de apagar todos os vetores
'If ubound(vet) >= 0 Then
'Response.Write "<br>" &"<a href='vetor.asp?action=LimpaVetor'>Apagar Tudo</a>" & "<br>" 'Imprime o Vetor na tela
'End if

End Function

Function LimpaVetor2

'Limpa todas as posi�oes do vetor, apagando a Session
session("GuardaVetor") = Empty

End Function




'///////////////////////////////////////// vetor Web Fam�lia /////////////////////////////////////////////////////////////////
Function VetorMonta3(Acao,Valor)
'Usamos o case para manipular a a��o da fun��o
Select Case Trim(Acao)
'Inclui nova posicao ao vetor
Case "Incluir"
'Guarda na variavel Vetor o conteudo da Session
Vetor = Session("GuardaVetor")
'Verifica se a Variavel Vetor � um Array, caso nao for entao definimos ela um Array
If Not IsArray(Vetor) Then Vetor = Array() End if
'Verifica se o Valor que esta sendo inserido j� esta no Vetor se estiver entao nao inseri para nao haver duplicidades do vetor
If InStr(Join(Vetor), Valor) = 0 Then
'Este comando ira preservar o vetor e adciona + 1 valor
ReDim preserve Vetor(UBound(Vetor)+1)
'Este � o valor que estamos adicionando no vetor
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
'Faz um la�o em todas as posi��es do vetor
For i = 0 To Ubound(AuxVetor)
'Verifica se o valor passado para excluir do vetor � diferente do valor que esta dentro da variavel Auxiliar
If AuxVetor(i) <> (Valor) Then
'Este comando ira preservar o vetor e adciona + 1 valor
ReDim preserve Vetor (UBound(Vetor)+1)
'Este � o valor que estamos adicionando no vetor
Vetor (Ubound(Vetor)) = AuxVetor(i)
'Coloca o conteudo da variavel vetor dentro da Session
Session("GuardaVetor") = Vetor
End If
Next
'Fim do Case
End Select
End Function

Function Incluir_Vetor3

'Executa a fun��o que ira criar uma posi��o do vetor, basta passar a acao e o valor
Call VetorMonta3("Incluir",Valor_Vetor)
'Request("Valor_Vetor")
'response.Write(Valor_Vetor&"=vet<BR>")
End Function


Function VisualizaValoresVetor3
vet = session("GuardaVetor")

'Veriofica se a Session � um array, caso nao for ent�o atribuimos a Session como um Array
IF Not IsArray(vet) Then vet = Array() End if
'Faremos um la�o entre todos os vetores criados

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
>Lista de Usu�rios</td>
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
'Verifica se a Session tem alguma posi��o, se tiver mostra a op��o de apagar todos os vetores
'If ubound(vet) >= 0 Then
'Response.Write "<br>" &"<a href='vetor.asp?action=LimpaVetor'>Apagar Tudo</a>" & "<br>" 'Imprime o Vetor na tela
'End if

End Function

Function LimpaVetor3

'Limpa todas as posi�oes do vetor, apagando a Session
session("GuardaVetor") = Empty

End Function


'///////////////////////////////////////// vetor alunos /////////////////////////////////////////////////////////////////
Function VetorMonta4(Acao,Valor)
'Usamos o case para manipular a a��o da fun��o
Select Case Trim(Acao)
'Inclui nova posicao ao vetor
Case "Incluir"
'Guarda na variavel Vetor o conteudo da Session
Vetor = Session("GuardaVetor")
'Verifica se a Variavel Vetor � um Array, caso nao for entao definimos ela um Array
If Not IsArray(Vetor) Then Vetor = Array() End if
'Verifica se o Valor que esta sendo inserido j� esta no Vetor se estiver entao nao inseri para nao haver duplicidades do vetor
If InStr(Join(Vetor), Valor) = 0 Then
'Este comando ira preservar o vetor e adciona + 1 valor
ReDim preserve Vetor(UBound(Vetor)+1)
'Este � o valor que estamos adicionando no vetor
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
'Faz um la�o em todas as posi��es do vetor
For i = 0 To Ubound(AuxVetor)
'Verifica se o valor passado para excluir do vetor � diferente do valor que esta dentro da variavel Auxiliar
If AuxVetor(i) <> (Valor) Then
'Este comando ira preservar o vetor e adciona + 1 valor
ReDim preserve Vetor (UBound(Vetor)+1)
'Este � o valor que estamos adicionando no vetor
Vetor (Ubound(Vetor)) = AuxVetor(i)
'Coloca o conteudo da variavel vetor dentro da Session
Session("GuardaVetor") = Vetor
End If
Next
'Fim do Case
End Select
End Function

Function Incluir_Vetor4

'Executa a fun��o que ira criar uma posi��o do vetor, basta passar a acao e o valor
Call VetorMonta("Incluir",Valor_Vetor)
'Request("Valor_Vetor")
'response.Write(Valor_Vetor&"=vet<BR>")
End Function


Function VisualizaValoresVetor4
vet = session("GuardaVetor")

'Veriofica se a Session � um array, caso nao for ent�o atribuimos a Session como um Array
IF Not IsArray(vet) Then vet = Array() End if
'Faremos um la�o entre todos os vetores criados

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
Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=ativos href =altera.asp?ori=01&cod_cons="&cod_cons&"&nvg="&chave&">"&no_aluno&"</a></font></li>")
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

response.Redirect("altera.asp?ori=01&cod_cons="&cod_cons&"&nvg="&chave)
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
        <form action="index.asp?opt=list&nvg=<%=chave%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
          <tr class="tb_tit"> 
            
      <td height="10" colspan="5">Preencha um dos campos abaixo</td>
          </tr>
          <TR>
		  
      <td height="26" valign="top"> 
        <table width="1000" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            
            <td width="147"  height="10"> 
              <div align="right"><font class="form_dado_texto"> Matr&iacute;cula:</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                </strong></font></div></td>
            
            <td width="62" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="busca1" type="text" class="textInput" id="busca1" size="12">
              </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font></td>
            
            <td width="147" height="10"> 
              <div align="right"><font class="form_dado_texto"> Nome: </font></div></td>
            
            <td width="392" height="10" ><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="busca2" type="text" class="textInput" id="busca2" size="55" maxlength="50">
              </font></td>
            
            <td width="250" height="10"><div align="center">
              <input name="Submit" type="submit" class="botao_prosseguir" id="Submit" value="Procurar">
              </div> </td>
          </tr>
		  </table>
		  </td>
		  </TR>
      </form>
      <tr>    
      	<td height="10"><hr> 
	 	</td>
  </tr>
<form name="alteracao" method="post" action="select_alunos.asp">      
      <tr>    
      	<td valign="top"> 
<table width="100%" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="250" class="tb_subtit"> 
        <div align="center">UNIDADE 
          </div></td>
      <td width="250" class="tb_subtit"> 
        <div align="center">CURSO 
          </div></td>
      <td width="250" class="tb_subtit"> 
        <div align="center">ETAPA 
          </div></td>
      <td width="250" class="tb_subtit"> 
        <div align="center">TURMA 
          </div></td>
      </tr>
    <tr> 
      <td width="250"> 
        <div align="center"> 
          <select name="unidade" class="select_style" id="unidade" onChange="recuperarCurso(this.value)">
            <option value="999990" selected></option>
            <%		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0
While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")
%>
            <option value="<%response.Write(NU_Unidade)%>"> 
              <%response.Write(NO_Abr)%>
              </option>
            <%RS0.MOVENEXT
WEND
%>
            </select>
          </div></td>
      <td width="250"> 
        <div align="center"> 
          <div id="divCurso"> 
            <select class="select_style">
              </select>
            </div>
          </div></td>
      <td width="250"> 
        <div align="center"> 
          <div id="divEtapa"> 
            <select class="select_style">
              </select>
            </div>
          </div></td>
      <td width="250"> 
        <div align="center"> 
          <div id="divTurma"> 
            <select class="select_style">
              </select>
            </div>
          </div></td>
      </tr>
    <tr>
      <td height="15" colspan="4" bgcolor="#FFFFFF"><hr></td>
      </tr>
<!--    <tr> 
      <td width="250" height="15" bgcolor="#FFFFFF"></td>
      <td width="250" height="15" bgcolor="#FFFFFF"></td>
      <td width="250" height="15" bgcolor="#FFFFFF"></td>
      <td width="250" height="15" bgcolor="#FFFFFF"><font size="2" face="Arial, Helvetica, sans-serif">
        <div align="center"><input name="Submit" type="submit" class="botao_prosseguir" id="Submit" value="Prosseguir"></div>
        </font></td>
      </tr>-->
    </table>        
	 	</td>
    </tr>  
  </FORM>         
      <tr>    
      	<td height="10">&nbsp; 
	 	</td>
  </tr>


        <%
end if
'Verifica se a Session tem alguma posi��o, se tiver mostra a op��o de apagar todos os vetores
'If ubound(vet) >= 0 Then
'Response.Write "<br>" &"<a href='vetor.asp?action=LimpaVetor'>Apagar Tudo</a>" & "<br>" 'Imprime o Vetor na tela
'End if

End Function

Function LimpaVetor4

'Limpa todas as posi�oes do vetor, apagando a Session
session("GuardaVetor") = Empty

End Function








Function contalunos (CAMINHOa,unidades,grau,serie,turma)
			
		Set CON_A = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_A.Open ABRIR
				
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidades &" AND CO_Curso = '"& grau &"' AND CO_Etapa = '"& serie &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)

conta=1
linha = 1
session("linha")= 0
if RS.EOF then

linha=0

else

while not RS.EOF
nu_chamada = RS("NU_Chamada")

if (conta = nu_chamada) then
linha=linha+1
conta=conta+1
else
falt_al = nu_chamada - conta
for k=1 to falt_al 
linha=linha+1
conta=conta+1
next

end if
  RS.MoveNext
Wend
end if
'response.write (linha)
session("linha")= linha
end function


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
			GeraNomesNovaVersao= "Mar�o"
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


Function GeraNomes(materia,unidades,grau,serie,Conexao)

Sqlmt= "SELECT * FROM TB_Materia where CO_Materia = '"& materia&"'"
Set rsmt= Conexao.Execute ( Sqlmt ) 
IF rsmt.eof THEN
no_materia= ""
ELSE
no_materia= rsmt("NO_Materia")
END IF


Sqlun= "SELECT * FROM TB_Unidade where NU_Unidade = "& unidades
Set rsun= Conexao.Execute ( Sqlun ) 
IF rsun.eof THEN
no_unidades= ""
ELSE
no_unidades= rsun("NO_Unidade")
END IF



Sqlgr= "SELECT * FROM TB_Curso where CO_Curso = '"& grau &"'"
Set rsgr= Conexao.Execute ( Sqlgr ) 
IF rsgr.eof THEN
no_grau= ""
ELSE
no_grau= rsgr("NO_Curso")
END IF


Sqlsr= "SELECT * FROM TB_Etapa where CO_Curso = '"& grau &"' and CO_Etapa = '"& serie &"'"
Set rssr= Conexao.Execute ( Sqlsr ) 
IF rssr.eof THEN
no_serie= ""
ELSE
no_serie= rssr("NO_Etapa")
END IF

session("no_materia") = no_materia
session("no_unidades") = no_unidades
session("no_grau") = no_grau
session("no_serie") = no_serie

end Function

Function GeraNomesMapao(unidades,grau,serie,Conexao)

Sqlun= "SELECT * FROM TB_Unidade where NU_Unidade = "& unidades
Set rsun= Conexao.Execute ( Sqlun ) 
no_unidades= rsun("NO_Unidade")

Sqlgr= "SELECT * FROM TB_Curso where CO_Curso = '"& grau &"'"
Set rsgr= Conexao.Execute ( Sqlgr ) 
no_grau= rsgr("NO_Curso")

Sqlsr= "SELECT * FROM TB_Etapa where CO_Curso = '"& grau &"' and CO_Etapa = '"& serie &"'"
Set rssr= Conexao.Execute ( Sqlsr ) 
no_serie= rssr("NO_Etapa")

session("no_materia") = no_materia
session("no_unidades") = no_unidades
session("no_grau") = no_grau
session("no_serie") = no_serie

end Function

'///////////////////////////////////////////////    �ltimo  //////////////////////////////////////////////////////////////

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
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' verifica se calcula m�dia ou n�o

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

elseif curso =1 and etapa=9 then
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
elseif curso =1 and etapa=99 then
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

'Function regra_aprovacao (unidade,curso,etapa,turma,total_nota,m1_aluno,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2)
'
'if total_nota>=2 then
'
'	Set CON0 = Server.CreateObject("ADODB.Connection") 
'	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
'	CON0.Open ABRIR0
'	
'	Set RSra = Server.CreateObject("ADODB.Recordset")
'	SQLra = "SELECT * FROM TB_Regras_Aprovacao where CO_Curso = '"&curso&"' and CO_Etapa = '"&etapa&"'"
'	RSra.Open SQLra, CON0	
'			
'	valor_m1=RSra("NU_Valor_Limite_M1")
'	res1_1=RSra("NO_Expr_Menor_M1")
'	res1_2=RSra("NO_Expr_Ma_Igual_M1")
'	peso_m2_m1=RSra("NU_Peso_Media_M2")
'	peso_m2_m2=RSra("NU_Peso_Media_M21")
'	
'	valor_m2=RSra("NU_Valor_Limite_M2")
'	res2_1=RSra("NO_Expr_Menor_M2")
'	res2_2=RSra("NO_Expr_Ma_Igual_M2")
'	peso_m3_m1=RSra("NU_Peso_Media_M3")
'	peso_m3_m2=RSra("NU_Peso_Media_M31")
'	peso_m3_m3=RSra("NU_Peso_Media_M32")
'	
'	valor_m3=RSra("NU_Valor_Limite_M3")
'	res3_1=RSra("NO_Expr_Menor_M3")
'	res3_2=RSra("NO_Expr_Ma_Igual_M3")
'
'		
'	m1_aluno=m1_aluno*1	
'	valor_m1=valor_m1*1
'	
'	if m1_aluno >= valor_m1 then
'	Session("resultado_1")=res1_2
'	else
'	Session("resultado_1")=res1_1
'	end if
'	if Session("resultado_1")=res1_2 then
'	else
'	
'				if nota_aux_m2_1="&nbsp;" then
'				m2_aluno="&nbsp;"
'				else								
'				m1_aluno_peso=m1_aluno*peso_m2_m1
'				nota_aux_m2_1_peso=nota_aux_m2_1*peso_m2_m2
'				m2_aluno=(m1_aluno_peso+nota_aux_m2_1_peso)/(peso_m2_m1+peso_m2_m2)
'							decimo = m2_aluno - Int(m2_aluno)
'								If decimo >= 0.5 Then
'									nota_arredondada = Int(m2_aluno) + 1
'									m2_aluno=nota_arredondada
'								Else
'									nota_arredondada = Int(m2_aluno)
'									m2_aluno=nota_arredondada					
'								End If
'							m2_aluno = formatNumber(m2_aluno,0)
'				end if
'		
'		if m2_aluno<>"&nbsp;" then		
'		m2_aluno=m2_aluno*1
'		valor_m2=valor_m2*1	
'		end if		
'		'response.write(m2_aluno &">="& valor_m2&"<BR>" )
'		if m2_aluno >= valor_m2 and nota_aux_m2_1<>"&nbsp;" then
'		Session("resultado_2")=res2_2
'		else
'		Session("resultado_2")=res2_1
'		end if
'
'		'if Session("resultado_2")=res2_2 or Session("resultado_1")=res1_2 then
'		'else
'		'		if nota_aux_m3_1="&nbsp;" then
'		'		m3_aluno="&nbsp;"
'		'		else
'		'		m1_aluno_peso=m1_aluno*peso_m3_m1					
'		'		m2_aluno_peso=m2_aluno*peso_m3_m2
'		'		nota_aux_m3_1_peso=nota_aux_m3_1*peso_m3_m3
'		'		m3_aluno=(m1_aluno_peso+m2_aluno_peso+nota_aux_m3_1_peso)/(peso_m3_m1+peso_m3_m2+peso_m3_m3)
'	'
'	'						decimo = m3_aluno - Int(m3_aluno)
'	'							If decimo >= 0.5 Then
'	'								nota_arredondada = Int(m3_aluno) + 1
'	'								m3_aluno=nota_arredondada
'	'							Else
'	'								nota_arredondada = Int(m3_aluno)
'	'								m3_aluno=nota_arredondada					
'	'							End If
'	'						m3_aluno = formatNumber(m3_aluno,0)
'	'			end if	
'	'		if m3_aluno<>"&nbsp;" then
'	'		m3_aluno=m3_aluno*1
'	'		valor_m3=valor_m3*1
'	'		end if
'	'		
'	'		if m3_aluno >= valor_m3 and nota_aux_m3_1<>"&nbsp;" then
'	'		Session("resultado_3")=res3_2
'	'		else
'	'		Session("resultado_3")=res3_1
'	'		end if
'	'	end if	
'	end if
'	Session("M2")=m2_aluno
'	Session("M3")=m3_aluno
'else
'	Session("resultado_1")="&nbsp;"
'	Session("resultado_2")="&nbsp;"
'	Session("resultado_3")="&nbsp;"
'	Session("M2")="&nbsp;"
'	Session("M3")="&nbsp;"	
'end if		
'end function
'Fun��o de Busca
'===================================================================================================
Function busca_por_nome(query,CONEXAO,tipo_busca)
'tipo_busca: alun=aluno, prof=professor
ano_letivo = session("ano_letivo") 

	'Converte caracteres que n�o s�o v�lidos em uma URL e os transformamem equivalentes para URL
	strProcura = Server.URLEncode(query)
	'Como nossa pesquisa ser� por "m�ltiplas palavras" (aqui voc� pode alterar ao seu gosto)
	'� necess�rio trocar o sinal de (=) pelo (%) que � usado com o LIKE na string SQL
	strProcura = replace(strProcura,"+"," ")
	strProcura = replace(strProcura,"%27","�")
	strProcura = replace(strProcura,"%27","'")
	strProcura = replace(strProcura,"%C0,","�")
	strProcura = replace(strProcura,"%C1","�")
	strProcura = replace(strProcura,"%C2","�")
	strProcura = replace(strProcura,"%C3","�")
	strProcura = replace(strProcura,"%C9","�")
	strProcura = replace(strProcura,"%CA","�")
	strProcura = replace(strProcura,"%CD","�")
	strProcura = replace(strProcura,"%D3","�")
	strProcura = replace(strProcura,"%D4","�")
	strProcura = replace(strProcura,"%D5","�")
	strProcura = replace(strProcura,"%DA","�")
	strProcura = replace(strProcura,"%DC","�")	
	strProcura = replace(strProcura,"%E1","�")
	strProcura = replace(strProcura,"%E1","�")
	strProcura = replace(strProcura,"%E2","�")
	strProcura = replace(strProcura,"%E3","�")
	strProcura = replace(strProcura,"%E7","�")
	strProcura = replace(strProcura,"%E9","�")
	strProcura = replace(strProcura,"%EA","�")
	strProcura = replace(strProcura,"%ED","�")
	strProcura = replace(strProcura,"%F3","�")
	strProcura = replace(strProcura,"F4","�")
	strProcura = replace(strProcura,"F5","�")
	strProcura = replace(strProcura,"%FA","�")
	strProcura = replace(strProcura,"%FC","�")

IF tipo_busca="alun" THEN
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Alunos.NO_Aluno like '%"& strProcura & "%' and TB_Matriculas.NU_Ano="&ano_letivo&" order BY NO_Aluno"
	'SQL = "SELECT * FROM TB_Alunos where NO_Aluno like '%"& strProcura & "%' order BY NO_Aluno"
	'response.Write(SQL)
	RS.Open SQL, CONEXAO		

	check_aluno=0
	WHile Not RS.EOF
		cod = RS("CO_Matricula")
		if check_aluno=0 then
			vetor_busca=cod		
		ELSE
			vetor_busca=vetor_busca&"#!#"&cod
		END IF
	check_aluno=check_aluno+1
	RS.MOVENEXT
	Wend
ELSEif tipo_busca="prof" THEN

		Set RS = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT * FROM TB_Professor where NO_Professor like '%"& strProcura & "%' order BY NO_Professor"
		RS.Open SQL, CONEXAO

	check_professor=0
	WHile Not RS.EOF
		cod = RS("CO_Professor")
		if check_professor=0 then
			vetor_busca=cod		
		ELSE
			vetor_busca=vetor_busca&"#!#"&cod
		END IF
	check_professor=check_professor+1
	RS.MOVENEXT
	Wend
END IF
busca_por_nome=vetor_busca	
End Function

%>