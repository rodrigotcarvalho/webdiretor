<!--#include file="../../../../../inc/caminhos.asp"-->
<%
Server.ScriptTimeout = 1200 'valor em segundos


opt=request.QueryString("opt")
if opt<>"f" then
	ano_letivo_wf=opt
	session("ano_letivo_upload")=ano_letivo_wf	
else
	ano_letivo_wf=request.QueryString("al")
	session("ano_letivo_upload")=ano_letivo_wf		
end if
	
vetor_tp_doc=request.QueryString("tp")
tp_doc_sel=request.QueryString("tp_sel")


ambiente_escola=request.QueryString("env")
	
if transicao = "S" then
	area="wd"
	link="http://simplynet2.tempsite.ws/wd/"&ambiente_escola
else
	if left(ambiente_escola,5)= "teste" then
		area="wdteste"
		link="http://www.simplynet.com.br/"&area&"/"&ambiente_escola
	else
		area="wd"
		link="http://www.webdiretor.com.br/"&ambiente_escola
	end if	
end if	

if opt ="f" then
'response.Write(session("tipo_arquivo_upl"))
'response.End()
	session("tipo_arquivo_upl")=session("tipo_arquivo_upl")*1
	if session("tipo_arquivo_upl")=0 then
		response.Redirect(link&"/wf/apoio/ctl/pub/upload.asp?opt=err1")
	else		
		tipo_arquivo=session("tipo_arquivo_upl")
						
	Set upl = Server.CreateObject("Persits.Upload")
	 %>
		 <!--#include file="connect_arquivo.asp"-->
	<% 
		'response.Write(caminho_pasta&session("tipo_arquivo_upl"))

		contarq = upl.Save(caminho_arquivo)
total = 0		
For Each File in upl.Files	
	nomeDoArquivo = split(File.Path,"\")	

	if total = 0 then
		file_nom = nomeDoArquivo(ubound(nomeDoArquivo))			
	else
		file_nom = file_nom&", "&nomeDoArquivo(ubound(nomeDoArquivo))		
	end if
total = total+1
next			

		
		if contarq = 0 then
			response.Redirect(link&"/wf/apoio/ctl/pub/upload.asp?opt=err")					
		end if
		

		
		Session("arquivos") = file_nom
							
		Session("upl_total") = upl.TotalBytes
	
		Session("ano_letivo_wf") =ano_letivo_wf				
		
		Dim objteste, objname, fso
		Dim nomeimagem 		
		set objteste = createobject("scripting.FileSystemObject")
		set objname = createobject("scripting.FileSystemObject")
		set fso = createobject("scripting.FileSystemObject")

		'Verifica se arquivo existe na pasta especifica
		'As funções abaixo tem como objetivo renomear os arquivos enviados
		if objteste.FileExists (caminho_arquivo& file1) = true then
			nom_doc =file1
			Set fileobject = fso.GetFile(caminho_arquivo& file1)
				'strReplacement = Server.URLEncode(nom_doc)
				strReplacement = replace(nom_doc,"+","_e_")
				strReplacement = replace(strReplacement," ","_")
				strReplacement = replace(strReplacement,"&","_e_")	
				strReplacement = replace(strReplacement,"-","")	
				strReplacement = replace(strReplacement,"–","_")																			
				strReplacement = replace(strReplacement,"´","")
				strReplacement = replace(strReplacement,"'","")
				strReplacement = replace(strReplacement,"Á","A")
				strReplacement = replace(strReplacement,"À","A")
				strReplacement = replace(strReplacement,"Â","A")
				strReplacement = replace(strReplacement,"Ã","A")
				strReplacement = replace(strReplacement,"Ç","C")			
				strReplacement = replace(strReplacement,"É","E")
				strReplacement = replace(strReplacement,"Ê","E")
				strReplacement = replace(strReplacement,"Í","I")
				strReplacement = replace(strReplacement,"Ó","O")
				strReplacement = replace(strReplacement,"Ô","O")
				strReplacement = replace(strReplacement,"Õ","O")
				strReplacement = replace(strReplacement,"Ú","U")
				strReplacement = replace(strReplacement,"Ü","U")	
				strReplacement = replace(strReplacement,"á","a")
				strReplacement = replace(strReplacement,"à","a")
				strReplacement = replace(strReplacement,"â","a")
				strReplacement = replace(strReplacement,"ã","a")
				strReplacement = replace(strReplacement,"ç","c")
				strReplacement = replace(strReplacement,"é","e")
				strReplacement = replace(strReplacement,"ê","e")
				strReplacement = replace(strReplacement,"í","i")
				strReplacement = replace(strReplacement,"ó","o")
				strReplacement = replace(strReplacement,"ô","o")
				strReplacement = replace(strReplacement,"õ","o")
				strReplacement = replace(strReplacement,"ú","u")
				strReplacement = replace(strReplacement,"ª","")
				strReplacement = replace(strReplacement,"º","")	
				strReplacement = replace(strReplacement,"Âº","")
				strReplacement = replace(strReplacement,"Ãº","")															
				nom_doc = replace(strReplacement,"ü","u")




			nomeimagem = nom_doc
			
			objteste.MoveFile caminho_arquivo& file1 ,caminho_arquivo&nomeimagem 

		else
			'Move arquivo para pasta específica
			'objteste.MoveFile "F:\teste\teste\"& doc &"\alt\"& teste &"","F:\teste\teste\"& doc &"\"

		end if 			
		
		if objteste.FileExists (caminho_pasta& file2) = true then
			nom_doc =file2
			Set fileobject = fso.GetFile(caminho_pasta& file2)
				'strReplacement = Server.URLEncode(nom_doc)
				strReplacement = replace(nom_doc,"+","_e_")
				strReplacement = replace(strReplacement," ","_")
				strReplacement = replace(strReplacement,"&","_e_")		
				strReplacement = replace(strReplacement,"-","")											
				strReplacement = replace(strReplacement,"´","")
				strReplacement = replace(strReplacement,"'","")
				strReplacement = replace(strReplacement,"Á","A")
				strReplacement = replace(strReplacement,"À","A")
				strReplacement = replace(strReplacement,"Â","A")
				strReplacement = replace(strReplacement,"Ã","A")
				strReplacement = replace(strReplacement,"Ç","C")			
				strReplacement = replace(strReplacement,"É","E")
				strReplacement = replace(strReplacement,"Ê","E")
				strReplacement = replace(strReplacement,"Í","I")
				strReplacement = replace(strReplacement,"Ó","O")
				strReplacement = replace(strReplacement,"Ô","O")
				strReplacement = replace(strReplacement,"Õ","O")
				strReplacement = replace(strReplacement,"Ú","U")
				strReplacement = replace(strReplacement,"Ü","U")	
				strReplacement = replace(strReplacement,"á","a")
				strReplacement = replace(strReplacement,"à","a")
				strReplacement = replace(strReplacement,"â","a")
				strReplacement = replace(strReplacement,"ã","a")
				strReplacement = replace(strReplacement,"ç","c")
				strReplacement = replace(strReplacement,"é","e")
				strReplacement = replace(strReplacement,"ê","e")
				strReplacement = replace(strReplacement,"í","i")
				strReplacement = replace(strReplacement,"ó","o")
				strReplacement = replace(strReplacement,"ô","o")
				strReplacement = replace(strReplacement,"õ","o")
				strReplacement = replace(strReplacement,"ú","u")
				strReplacement = replace(strReplacement,"ª","")
				strReplacement = replace(strReplacement,"º","")
				strReplacement = replace(strReplacement,"Âº","")
				strReplacement = replace(strReplacement,"Ãº","")													
				nom_doc = replace(strReplacement,"ü","u")

			nomeimagem = nom_doc
			
			objteste.MoveFile caminho_pasta& file2 ,caminho_pasta&session("tipo_arquivo_upl")&"\"&nomeimagem 

		else
			'Move arquivo para pasta específica
			'objteste.MoveFile "F:\teste\teste\"& doc &"\alt\"& teste &"","F:\teste\teste\"& doc &"\"

		end if 					
		
		if objteste.FileExists (caminho_pasta& file3) = true then
			nom_doc =file3
			Set fileobject = fso.GetFile(caminho_pasta& file3)
				'strReplacement = Server.URLEncode(nom_doc)
				strReplacement = replace(nom_doc,"+","_e_")
				strReplacement = replace(strReplacement," ","_")
				strReplacement = replace(strReplacement,"&","_e_")		
				strReplacement = replace(strReplacement,"-","")											
				strReplacement = replace(strReplacement,"´","")
				strReplacement = replace(strReplacement,"'","")
				strReplacement = replace(strReplacement,"Á","A")
				strReplacement = replace(strReplacement,"À","A")
				strReplacement = replace(strReplacement,"Â","A")
				strReplacement = replace(strReplacement,"Ã","A")
				strReplacement = replace(strReplacement,"Ç","C")			
				strReplacement = replace(strReplacement,"É","E")
				strReplacement = replace(strReplacement,"Ê","E")
				strReplacement = replace(strReplacement,"Í","I")
				strReplacement = replace(strReplacement,"Ó","O")
				strReplacement = replace(strReplacement,"Ô","O")
				strReplacement = replace(strReplacement,"Õ","O")
				strReplacement = replace(strReplacement,"Ú","U")
				strReplacement = replace(strReplacement,"Ü","U")	
				strReplacement = replace(strReplacement,"á","a")
				strReplacement = replace(strReplacement,"à","a")
				strReplacement = replace(strReplacement,"â","a")
				strReplacement = replace(strReplacement,"ã","a")
				strReplacement = replace(strReplacement,"ç","c")
				strReplacement = replace(strReplacement,"é","e")
				strReplacement = replace(strReplacement,"ê","e")
				strReplacement = replace(strReplacement,"í","i")
				strReplacement = replace(strReplacement,"ó","o")
				strReplacement = replace(strReplacement,"ô","o")
				strReplacement = replace(strReplacement,"õ","o")
				strReplacement = replace(strReplacement,"ú","u")
				strReplacement = replace(strReplacement,"ª","")
				strReplacement = replace(strReplacement,"º","")
				strReplacement = replace(strReplacement,"Âº","")	
				strReplacement = replace(strReplacement,"Ãº","")												
				nom_doc = replace(strReplacement,"ü","u")


			nomeimagem = nom_doc
			
			objteste.MoveFile caminho_pasta& file3 ,caminho_pasta&session("tipo_arquivo_upl")&"\"&nomeimagem 

		else
			'Move arquivo para pasta específica
			'objteste.MoveFile "F:\teste\teste\"& doc &"\alt\"& teste &"","F:\teste\teste\"& doc &"\"
		end if 					
		
		if objteste.FileExists (caminho_pasta& file4) = true then
			nom_doc =file4
			Set fileobject = fso.GetFile(caminho_pasta& file4)
				'strReplacement = Server.URLEncode(nom_doc)
				strReplacement = replace(nom_doc,"+","_e_")
				strReplacement = replace(strReplacement," ","_")
				strReplacement = replace(strReplacement,"&","_e_")	
				strReplacement = replace(strReplacement,"-","")												
				strReplacement = replace(strReplacement,"´","")
				strReplacement = replace(strReplacement,"'","")
				strReplacement = replace(strReplacement,"Á","A")
				strReplacement = replace(strReplacement,"À","A")
				strReplacement = replace(strReplacement,"Â","A")
				strReplacement = replace(strReplacement,"Ã","A")
				strReplacement = replace(strReplacement,"Ç","C")			
				strReplacement = replace(strReplacement,"É","E")
				strReplacement = replace(strReplacement,"Ê","E")
				strReplacement = replace(strReplacement,"Í","I")
				strReplacement = replace(strReplacement,"Ó","O")
				strReplacement = replace(strReplacement,"Ô","O")
				strReplacement = replace(strReplacement,"Õ","O")
				strReplacement = replace(strReplacement,"Ú","U")
				strReplacement = replace(strReplacement,"Ü","U")	
				strReplacement = replace(strReplacement,"á","a")
				strReplacement = replace(strReplacement,"à","a")
				strReplacement = replace(strReplacement,"â","a")
				strReplacement = replace(strReplacement,"ã","a")
				strReplacement = replace(strReplacement,"ç","c")
				strReplacement = replace(strReplacement,"é","e")
				strReplacement = replace(strReplacement,"ê","e")
				strReplacement = replace(strReplacement,"í","i")
				strReplacement = replace(strReplacement,"ó","o")
				strReplacement = replace(strReplacement,"ô","o")
				strReplacement = replace(strReplacement,"õ","o")
				strReplacement = replace(strReplacement,"ú","u")
				strReplacement = replace(strReplacement,"ª","")
				strReplacement = replace(strReplacement,"º","")
				strReplacement = replace(strReplacement,"Âº","")	
				strReplacement = replace(strReplacement,"Ãº","")												
				nom_doc = replace(strReplacement,"ü","u")


			nomeimagem = nom_doc
			
			objteste.MoveFile caminho_pasta& file4 ,caminho_pasta&session("tipo_arquivo_upl")&"\"&nomeimagem 

		else
			'Move arquivo para pasta específica
			'objteste.MoveFile "F:\teste\teste\"& doc &"\alt\"& teste &"","F:\teste\teste\"& doc &"\"

		end if 		
		
		if objteste.FileExists (caminho_pasta& file5) = true then
			nom_doc =file5
			Set fileobject = fso.GetFile(caminho_pasta& file5)
				'strReplacement = Server.URLEncode(nom_doc)
				strReplacement = replace(nom_doc,"+","_e_")
				strReplacement = replace(strReplacement," ","_")
				strReplacement = replace(strReplacement,"&","_e_")		
				strReplacement = replace(strReplacement,"-","")											
				strReplacement = replace(strReplacement,"´","")
				strReplacement = replace(strReplacement,"'","")
				strReplacement = replace(strReplacement,"Á","A")
				strReplacement = replace(strReplacement,"À","A")
				strReplacement = replace(strReplacement,"Â","A")
				strReplacement = replace(strReplacement,"Ã","A")
				strReplacement = replace(strReplacement,"Ç","C")			
				strReplacement = replace(strReplacement,"É","E")
				strReplacement = replace(strReplacement,"Ê","E")
				strReplacement = replace(strReplacement,"Í","I")
				strReplacement = replace(strReplacement,"Ó","O")
				strReplacement = replace(strReplacement,"Ô","O")
				strReplacement = replace(strReplacement,"Õ","O")
				strReplacement = replace(strReplacement,"Ú","U")
				strReplacement = replace(strReplacement,"Ü","U")	
				strReplacement = replace(strReplacement,"á","a")
				strReplacement = replace(strReplacement,"à","a")
				strReplacement = replace(strReplacement,"â","a")
				strReplacement = replace(strReplacement,"ã","a")
				strReplacement = replace(strReplacement,"ç","c")
				strReplacement = replace(strReplacement,"é","e")
				strReplacement = replace(strReplacement,"ê","e")
				strReplacement = replace(strReplacement,"í","i")
				strReplacement = replace(strReplacement,"ó","o")
				strReplacement = replace(strReplacement,"ô","o")
				strReplacement = replace(strReplacement,"õ","o")
				strReplacement = replace(strReplacement,"ú","u")
				strReplacement = replace(strReplacement,"ª","")
				strReplacement = replace(strReplacement,"º","")
				strReplacement = replace(strReplacement,"Âº","")	
				strReplacement = replace(strReplacement,"Ãº","")																	
				nom_doc = replace(strReplacement,"ü","u")


			nomeimagem = nom_doc
			
			objteste.MoveFile caminho_pasta& file5 ,caminho_pasta&session("tipo_arquivo_upl")&"\"&nomeimagem 

		else
			'Move arquivo para pasta específica
			'objteste.MoveFile "F:\teste\teste\"& doc &"\alt\"& teste &"","F:\teste\teste\"& doc &"\"

		end if 								
		
		response.Redirect("criarquivo.asp?opt=i&env="&ambiente_escola)					
		Set upl = Nothing 
		set fileObject = nothing
		set objteste = nothing
		set objname = nothing
		set fso = nothing 			
						
	end if
end if	
%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
<!--
// A função abaixo pega a versão mais nova do xmlhttp do IE e verifica se é Firefox. Funciona nos dois.
function createXMLHTTP()
            {
                        try
                        {
                                   ajax = new ActiveXObject("Microsoft.XMLHTTP");
                        }
                        catch(e)
                        {
                                   try
                                   {
                                               ajax = new ActiveXObject("Msxml2.XMLHTTP");
                                               alert(ajax);
                                   }
                                   catch(ex)
                                   {
                                               try
                                               {
                                                           ajax = new XMLHttpRequest();
                                               }
                                               catch(exc)
                                               {
                                                            alert("Esse browser não tem recursos para uso do Ajax");
                                                            ajax = null;
                                               }
                                   }
                                   return ajax;
                        }
           
           
               var arrSignatures = ["MSXML2.XMLHTTP.5.0", "MSXML2.XMLHTTP.4.0",
               "MSXML2.XMLHTTP.3.0", "MSXML2.XMLHTTP",
               "Microsoft.XMLHTTP"];
               for (var i=0; i < arrSignatures.length; i++) {
                                                                          try {
                                                                                                             var oRequest = new ActiveXObject(arrSignatures[i]);
                                                                                                             return oRequest;
                                                                          } catch (oError) {
                                                                          }
                                      }
           
                                      throw new Error("MSXML is not installed on your system.");
                        }                                
						
						
						 function gravarTipo(tpTipo)
                                   {
// Criação do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicitação HTTP. O primeiro parâmetro informa o método post/get
// O segundo parâmetro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicitação síncrona, o parâmetro deve ser false
                                               oHTTPRequest.open("post", "grava.asp?opt=t", true);
// Para solicitações utilizando o método post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A função abaixo é executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto já completou a solicitação
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto é gerado no arquivo executa.asp e colocado no div
                                                                       document.all.divCurso.innerHTML = oHTTPRequest.responseText;
                                                           }
                                               }
// Abaixo é enviada a solicitação. Note que a configuração
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("tp_pub=" + tpTipo);
                                   }

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>							   
<link href="estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0">
<table width="1000" height="170" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
    <td width="1000" valign="top"> <table width="1000" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
        <tr> 
          <td width="1000"><FORM METHOD="POST" ENCTYPE="multipart/form-data" ACTION="envia.asp?opt=f&al=<%Response.Write(ano_letivo_wf)%>&env=<%Response.Write(ambiente_escola)%>" target="_parent">
              <table width="1000" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
                <tr bgcolor="#FFFFFF" background="../../../../img/fundo_interno.gif"> 
                  <td height="20"> <div align="right"><font class="form_dado_texto"> 
                  Nome da Pasta de Documentos: </font></div></td>
                  <td height="20"> <select name="tipo_doc" class="borda" onChange="gravarTipo(this.value)">
                                    <%
			tipos_doc=split(vetor_tp_doc,"$!$")
			for d=0 to ubound(tipos_doc)
			
				dados_tipos_doc=split(tipos_doc(d),"!$!")			
				cod_tp_doc=dados_tipos_doc(0)
				nom_tp_doc=dados_tipos_doc(1)	
				strReplacement = Server.URLEncode(nom_tp_doc)
				strReplacement = replace(strReplacement,"+"," ")
				strReplacement = replace(strReplacement,"%27","´")
				strReplacement = replace(strReplacement,"%B4","'")
				strReplacement = replace(strReplacement,"%C0,","Á")
				strReplacement = replace(strReplacement,"%C1","À")
				strReplacement = replace(strReplacement,"%C2","Â")
				strReplacement = replace(strReplacement,"%C3","Ã")
				strReplacement = replace(strReplacement,"%C9","É")
				strReplacement = replace(strReplacement,"%CA","Ê")
				strReplacement = replace(strReplacement,"%CD","Í")
				strReplacement = replace(strReplacement,"%D3","Ó")
				strReplacement = replace(strReplacement,"%D4","Ô")
				strReplacement = replace(strReplacement,"%D5","Õ")
				strReplacement = replace(strReplacement,"%DA","Ú")
				strReplacement = replace(strReplacement,"%DC","Ü")	
				strReplacement = replace(strReplacement,"%E0","á")
				strReplacement = replace(strReplacement,"%E1","à")
				strReplacement = replace(strReplacement,"%E2","â")
				strReplacement = replace(strReplacement,"%E3","ã")
				strReplacement = replace(strReplacement,"%E7","ç")
				strReplacement = replace(strReplacement,"%E9","é")
				strReplacement = replace(strReplacement,"%EA","ê")
				strReplacement = replace(strReplacement,"%ED","í")
				strReplacement = replace(strReplacement,"%F3","ó")
				strReplacement = replace(strReplacement,"%F4","ô")
				strReplacement = replace(strReplacement,"%F5","õ")
				strReplacement = replace(strReplacement,"%FA","ú")
				nom_tp_doc = replace(strReplacement,"%FC","ü")
				
				d=d*1
				if tp_doc_sel="" or isnull(tp_doc_sel) then
					if d=0 then
						session("tipo_arquivo_upl")=cod_tp_doc
					end if
				else
					session("tipo_arquivo_upl")=tp_doc_sel
				end if		
					
				if cod_tp_doc=session("tipo_arquivo_upl") then
					selected="SELECTED"
				else	
					selected=""
				end if						
%>
                <option value="<%response.Write(cod_tp_doc)%>" <%response.Write(selected)%>>
                <%response.Write(nom_tp_doc)%>
                </option>
                <%
			Next
%>
                    </select> </td>
                </tr>
                <tr> 
                  <td  width="350"> <div align="right"><font class="form_dado_texto">Arquivo 
                      1: </font></div></td>
                  <td width="650"> <INPUT TYPE=FILE SIZE=60 NAME="FILE1" class="borda"></td>
                </tr>
                <tr> 
                  <td width="350"> <div align="right"><font class="form_dado_texto">Arquivo 
                      2: </font></div></td>
                  <td width="650"> <INPUT TYPE=FILE SIZE=60 NAME="FILE2" class="borda"></td>
                </tr>
                <tr> 
                  <td width="350"> <div align="right"><font class="form_dado_texto">Arquivo 
                      3: </font></div></td>
                  <td width="650"> <INPUT TYPE=FILE SIZE=60 NAME="FILE3" class="borda"></td>
                </tr>
                <tr> 
                  <td width="350"> <div align="right"><font class="form_dado_texto">Arquivo 
                      4: </font></div></td>
                  <td width="650"> <INPUT TYPE=FILE SIZE=60 NAME="FILE4" class="borda"></td>
                </tr>
                <tr> 
                  <td width="350"> <div align="right"><font class="form_dado_texto">Arquivo 
                      5: </font></div></td>
                  <td width="650"> <INPUT TYPE=FILE SIZE=60 NAME="FILE5" class="borda"></td>
                </tr>
                <tr> 
                  <td colspan="2"><hr width="1000"></td>
                </tr>
                <tr> 
                  <td colspan="2"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="33%"> <div align="center"> 
                            <input name="SUBMIT5" type=button class="botao_cancelar" onClick="MM_goToURL('parent','<%response.Write(link)%>/wf/apoio/ctl/pub/docs.asp?opt=f&pagina=1&v=s');return document.MM_returnValue" value="Voltar">
                        </div></td>
                        <td width="34%"> <div align="center" id="divCurso"></div></td>
                        <td width="33%"> <div align="center"> 
                            <input name="SUBMIT" type=SUBMIT class="botao_prosseguir" value="Upload!">
							
                        </div></td>
                      </tr>
                    </table></td>
                </tr>
              </table>
            </form></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>