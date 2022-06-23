<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%
opt = request.QueryString("opt")

ano_letivo_wf = Session("ano_letivo_wf")
co_usr = session("co_user")
tipo_arquivo = request.Form("pasta")
unidade = request.Form("unidade")
curso = request.Form("curso")
etapa = request.Form("etapa")
turma = request.Form("turma")

unidade = unidade*1
'if unidade <> 999990 then
'	if isnull(curso)= false then
'	   curso = curso*1	
'	   if curso<>999990 then
'			if isnull(etapa)= false then
'				if curso>0 then
'				   etapa = etapa*1
'				   if etapa<>999990 then
'						if isnull(turma)= false then
'						   turma = turma*1	
'						   if turma<>999990 then				   
'								query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"' AND Turma='"&turma&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"' AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null) "
'							'	query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"' AND Turma='"&turma&"') "
'							else
'								query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa Is Null AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null) "	
'							'	query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') "																			
'							end if
'						else
'							query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa Is Null AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null) "	
'						'	query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') "							
'						end if
'					else
'						query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null) "		
'	'					query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"') "																					
'					end if
'				else
'				   if etapa<>"999990" then
'						if isnull(turma)= false then
'						   turma = turma*1	
'						   if turma<>999990 then				   
'								query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"' AND Turma='"&turma&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"' AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null) "
'							'	query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"' AND Turma='"&turma&"') "
'							else
'								query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa Is Null AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null) "	
'							'	query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') "																			
'							end if
'						else
'							query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa Is Null AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null) "	
'						'	query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') "							
'						end if
'					else
'						query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null) "		
'	'					query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"') "																					
'					end if				
'				end if	
'			else
'				query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null) "		
'				'query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"') "					
'			end if			
'		else			
'			query= " AND (Unidade='"&unidade&"') OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null) "	
'			'query= " AND (Unidade='"&unidade&"') "								
'		end if
'	else
'			query= " AND (Unidade='"&unidade&"') OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null) "	
'			'query= " AND (Unidade='"&unidade&"') "			
'	end if	
'else
'	query=" "		
'end if			
unidade = unidade*1
if unidade <> 999990 then
	if isnull(curso)= false then
	   curso = curso*1	
	   if curso<>999990 then
			if isnull(etapa)= false then
				if curso>0 then
				   etapa = etapa*1
				   if etapa<>999990 then
						if isnull(turma)= false then
							if isnumeric(turma) then
						  	 turma = turma*1
							   if turma<>999990 then				   
									query= " AND ((Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"' AND Turma='"&turma&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"' AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null)) "
								'	query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"' AND Turma='"&turma&"') "
								else
									query= " AND ((Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa Is Null AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null)) "	
								'	query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') "																			
								end if		
							else
							   if turma<>"999990" then				   
									query= " AND ((Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"' AND Turma='"&turma&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"' AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null)) "
								'	query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"' AND Turma='"&turma&"') "
								else
									query= " AND ((Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa Is Null AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null)) "	
								'	query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') "																			
								end if													 	
							end if 
						else
							query= " AND ((Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa Is Null AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null)) "	
						'	query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') "							
						end if
					else
						query= " AND ((Unidade='"&unidade&"' AND Curso='"&curso&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null)) "		
	'					query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"') "																					
					end if
				else
				   if etapa<>"999990" then
						if isnull(turma)= false then
							if isnumeric(turma) then
						  	 turma = turma*1	
							   if turma<>999990 then				   
									query= " AND ((Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"' AND Turma='"&turma&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"' AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null)) "
								'	query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"' AND Turma='"&turma&"') "
								else
									query= " AND ((Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa Is Null AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null)) "	
								'	query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') "																			
								end if
							else
							   if turma<>"999990" then				   
									query= " AND ((Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"' AND Turma='"&turma&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"' AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null)) "
								'	query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"' AND Turma='"&turma&"') "
								else
									query= " AND ((Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa Is Null AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null)) "	
								'	query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') "																			
								end if							
							end if	
						else
							query= " AND ((Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa Is Null AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null)) "	
						'	query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"' AND  Etapa='"&etapa&"') "							
						end if
					else
						query= " AND ((Unidade='"&unidade&"' AND Curso='"&curso&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null)) "		
	'					query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"') "																					
					end if				
				end if	
			else
				query= " AND ((Unidade='"&unidade&"' AND Curso='"&curso&"') OR (Unidade='"&unidade&"' AND Curso Is Null AND  Etapa Is Null AND Turma Is Null) OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null)) "		
				'query= " AND (Unidade='"&unidade&"' AND Curso='"&curso&"') "					
			end if			
		else			
			query= " AND ((Unidade='"&unidade&"') OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null)) "	
			'query= " AND (Unidade='"&unidade&"') "								
		end if
	else
			query= " AND ((Unidade='"&unidade&"') OR (Unidade Is Null AND Curso Is Null AND Etapa Is Null AND Turma Is Null)) "	
			'query= " AND (Unidade='"&unidade&"') "			
	end if	
else
	query=" "		
end if			

nivel=4
		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		

    	Set CON_WF = Server.CreateObject("ADODB.Connection") 
		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_WF.Open ABRIR_WF




chave=session("chave")
session("chave")=chave
nvg_split=split(chave,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano = DatePart("yyyy", now) 
mes = DatePart("m", now) 
dia = DatePart("d", now) 


if ano_letivo_wf=ano then
	data_expira=dia&"/"&mes&"/"&ano
else

	data_expira="31/12/"&ano_letivo_wf
end if	

ano_info=nivel&"-"&chave&"-"&ano_letivo_wf


call navegacao (CON,chave,nivel)
navega=Session("caminho")	

	Set RS2 = Server.CreateObject("ADODB.Recordset")
	SQL2 = "SELECT * FROM TB_Autoriz_Usuario_Grupo Where CO_Usuario = "&co_usr
	RS2.Open SQL2, CON
		
if RS2.EOF then

else		
	co_grupo=RS2("CO_Grupo")
End if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="file:../../../../img/mm_menu.js"></script>
<script type="text/javascript" src="../../../../js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_nbGroup(event, grpName) { //v6.0
  var i,img,nbArr,args=MM_nbGroup.arguments;
  if (event == "init" && args.length > 2) {
    if ((img = MM_findObj(args[2])) != null && !img.MM_init) {
      img.MM_init = true; img.MM_up = args[3]; img.MM_dn = img.src;
      if ((nbArr = document[grpName]) == null) nbArr = document[grpName] = new Array();
      nbArr[nbArr.length] = img;
      for (i=4; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
        if (!img.MM_up) img.MM_up = img.src;
        img.src = img.MM_dn = args[i+1];
        nbArr[nbArr.length] = img;
    } }
  } else if (event == "over") {
    document.MM_nbOver = nbArr = new Array();
    for (i=1; i < args.length-1; i+=3) if ((img = MM_findObj(args[i])) != null) {
      if (!img.MM_up) img.MM_up = img.src;
      img.src = (img.MM_dn && args[i+2]) ? args[i+2] : ((args[i+1])? args[i+1] : img.MM_up);
      nbArr[nbArr.length] = img;
    }
  } else if (event == "out" ) {
    for (i=0; i < document.MM_nbOver.length; i++) {
      img = document.MM_nbOver[i]; img.src = (img.MM_dn) ? img.MM_dn : img.MM_up; }
  } else if (event == "down") {
    nbArr = document[grpName];
    if (nbArr)
      for (i=0; i < nbArr.length; i++) { img=nbArr[i]; img.src = img.MM_up; img.MM_dn = 0; }
    document[grpName] = nbArr = new Array();
    for (i=2; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
      if (!img.MM_up) img.MM_up = img.src;
      img.src = img.MM_dn = (args[i+1])? args[i+1] : img.MM_up;
      nbArr[nbArr.length] = img;
  } }
}

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
function submitano()  
{
   var f=document.forms[0]; 
      f.submit(); 
}
function submitsistema()  
{
   var f=document.forms[1]; 
      f.submit(); 
}
function submitrapido()  
{
   var f=document.forms[2]; 
      f.submit(); 
}  
function submitforminterno()  
{
   var f=document.forms[3]; 
      f.submit(); 
	  
}
function checksubmit()
{
 if (document.form.pasta.value == "nulo")
  {    alert("Por favor selecione uma pasta!")
   document.form.pasta.focus()
    return false
 }
 
  return true
}
//-->
</script>
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
						
						
						 function recuperarCurso(uTipo)
                                   {
// Criação do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicitação HTTP. O primeiro parâmetro informa o método post/get
// O segundo parâmetro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicitação síncrona, o parâmetro deve ser false
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=c", true);
// Para solicitações utilizando o método post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A função abaixo é executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto já completou a solicitação
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto é gerado no arquivo executa.asp e colocado no div
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
document.all.divEtapa.innerHTML ="<select class=borda></select>"
document.all.divTurma.innerHTML = "<select class=borda></select>"
//recuperarEtapa()
                                                           }
                                               }
// Abaixo é enviada a solicitação. Note que a configuração
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {
// Criação do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicitação HTTP. O primeiro parâmetro informa o método post/get
// O segundo parâmetro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicitação síncrona, o parâmetro deve ser false
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=e", true);
// Para solicitações utilizando o método post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A função abaixo é executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto já completou a solicitação
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto é gerado no arquivo executa.asp e colocado no div
                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select class=borda></select>"
//recuperarTurma()
                                                           }
                                               }
// Abaixo é enviada a solicitação. Note que a configuração
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {
// Criação do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicitação HTTP. O primeiro parâmetro informa o método post/get
// O segundo parâmetro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicitação síncrona, o parâmetro deve ser false
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t", true);
// Para solicitações utilizando o método post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A função abaixo é executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto já completou a solicitação
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto é gerado no arquivo executa.asp e colocado no div
                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t																	   
                                                           }
                                               }
// Abaixo é enviada a solicitação. Note que a configuração
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("e_pub=" + eTipo);

                                   }

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0">
<% call cabecalho (nivel)
	  %>
	<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
		<tr>
			<td height="10" class="tb_caminho"><font class="style-caminho">
				<%
	  response.Write(navega)

%>
				</font></td>
		</tr>
		<%
if opt = "a" then%>
		<tr>
			<td height="10"><%
		call mensagens(nivel,9705,2,0)
%></td>
		</tr>
		<% 	end if 

%>
		<tr>
			<td height="10"><%	call mensagens(nivel,9704,0,0) 
	  				  
%></td>
		</tr>
		<tr>		
		<td valign="top"><form action="docs.asp" method="post" name="form" onSubmit="return checksubmit()"><table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
			<tr>
				<td valign="top" class="tb_tit">Verificar Documentos Publicados</td>
			</tr>
			<tr>
				<td valign="top">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td>
							<table width="100%" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td width="200" class="tb_subtit"><div align="center">UNIDADE </div></td>
								<td width="200" class="tb_subtit"><div align="center">CURSO </div></td>
								<td width="200" class="tb_subtit"><div align="center">ETAPA </div></td>
								<td width="200" class="tb_subtit"><div align="center">TURMA </div></td>
								<td width="200" class="tb_subtit"><div align="center">PASTAS PUBLICADAS</div></td>
							</tr>
							<tr>
								<td width="200"><div align="center">
									<select name="unidade" class="select_style" onChange="recuperarCurso(this.value)">	
<% if isnull(unidade)=false then
		unidade=unidade*1
		if unidade =999990 then
			response.Write("<option value=""999990"" selected></option>")
		else
			response.Write("<option value=""999990""></option>")	
		end if	
	else
		response.Write("<option value=""999990""selected></option>")
	end if			
	
			Set RS0 = Server.CreateObject("ADODB.Recordset")
			SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
			RS0.Open SQL0, CON0
While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")
unidade=unidade*1
NU_Unidade=NU_Unidade*1
if NU_Unidade=unidade then
%>
										<option value="<%response.Write(NU_Unidade)%>" selected>
											<%response.Write(NO_Abr)%>
											</option>
										<%
else
%>
										<option value="<%response.Write(NU_Unidade)%>">
											<%response.Write(NO_Abr)%>
											</option>
										<%
end if
RS0.MOVENEXT
WEND
%>
									</select>
								</div></td>
								<td width="200"><div align="center">
									<div id="divCurso">
										<select name="curso" class="select_style" onChange="recuperarEtapa(this.value)">
											<%	
	if isnull(curso)=false then
		curso=curso*1
		if curso =999990 then
			response.Write("<option value=""999990"" selected></option>")
		else
			response.Write("<option value=""999990""></option>")	
		end if	
	else
		response.Write("<option value=""999990""selected></option>")
	end if												
												
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT Distinct CO_Curso FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade
		RS0.Open SQL0, CON0
		
While not RS0.EOF
CO_Curso = RS0("CO_Curso")

		Set RS0a = Server.CreateObject("ADODB.Recordset")
		SQL0a = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
		RS0a.Open SQL0a, CON0
		
NO_Curso = RS0a("NO_Abreviado_Curso")		
curso=curso*1
CO_Curso = CO_Curso*1
if CO_Curso=curso then
%>
											<option value="<%response.Write(CO_Curso)%>" selected>
												<%response.Write(NO_Curso)%>
												</option>
											<%
else
%>
											<option value="<%response.Write(CO_Curso)%>">
												<%response.Write(NO_Curso)%>
												</option>
											<%
end if
RS0.MOVENEXT
WEND
%>
										</select>
									</div>
								</div></td>
								<td width="200"><div align="center">
									<div id="divEtapa">
										<select name="etapa" class="select_style" onChange="recuperarTurma(this.value)">
											<%		

	if isnull(etapa)=false then
		if curso>0 then
			etapa=etapa*1
			if etapa =999990 then
				response.Write("<option value=""999990"" selected></option>")
			else
				response.Write("<option value=""999990""></option>")	
			end if	
		else
			if etapa ="999990" then
				response.Write("<option value=""999990"" selected></option>")
			else
				response.Write("<option value=""999990""></option>")	
			end if			
		end if	
	else
		response.Write("<option value=""999990""selected></option>")
	end if	

		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"'"
		RS0b.Open SQL0b, CON0
		
		
While not RS0b.EOF
co_etapa = RS0b("CO_Etapa")


		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&curso&"' AND CO_Etapa='"&co_etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
if isnumeric(etapa) then
	etapa=etapa*1
end if	
if isnumeric(co_etapa) then
	co_etapa=co_etapa*1
end if	

if co_etapa=etapa then
%>
											<option value="<%response.Write(co_etapa)%>" selected>
												<%response.Write(NO_Etapa)%>
												</option>
											<%
else
%>
											<option value="<%response.Write(co_etapa)%>">
												<%response.Write(NO_Etapa)%>
												</option>
											<%

end if
RS0b.MOVENEXT
WEND
%>
										</select>
									</div>
								</div></td>
								<td width="200"><div align="center">
									<div id="divTurma">
										<select name="turma" class="select_style" onChange="MM_callJS('submitfuncao()')">
											<%
	if isnull(turma)=false then
		if isnumeric(turma) then
		  turma=turma*1
			if turma =999990 then
				response.Write("<option value=""999990"" selected></option>")
			else
				response.Write("<option value=""999990""></option>")	
			end if	
		else
			if turma ="999990" then
				response.Write("<option value=""999990"" selected></option>")
			else
				response.Write("<option value=""999990""></option>")	
			end if				
		end if	
	else
		response.Write("<option value=""999990""selected></option>")
	end if												
											
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Turma where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"' AND CO_Etapa='" & etapa & "' order by CO_Turma" 
		RS3.Open SQL3, CON0						

while not RS3.EOF
co_turma= RS3("CO_Turma")
if isnumeric(turma) then
turma=turma*1
end if
if isnumeric(co_turma) then
co_turma=co_turma*1
end if
if co_turma=turma then
%>
											<option value="<%response.Write(co_turma)%>" selected>
												<%response.Write(co_turma)%>
												</option>
											<%
else
%>
											<option value="<%=co_turma%>">
												<%response.Write(co_turma)%>
												</option>
											<%
co_turma_check = co_turma
end if
RS3.MOVENEXT
WEND
%>
										</select>
									</div>
								</div></td>
								<td width="200"><div align="center">
									<%
		
		Set RSt = Server.CreateObject("ADODB.Recordset")
		SQLt = "SELECT COUNT(CO_Pasta_Doc) AS TOTAL_PASTAS FROM TB_Tipo_Pasta_Doc where ((DA_Expira NOT BETWEEN #01/01/1900# AND #"&data_expira&"#) AND IN_Expira= TRUE) or IN_Expira= FALSE"
		RST.Open SQLt, CON0
		
		Set RS_doc = Server.CreateObject("ADODB.Recordset")
		SQL_doc = "SELECT CO_Pasta_Doc, NO_Pasta FROM TB_Tipo_Pasta_Doc where ((DA_Expira NOT BETWEEN #01/01/1900# AND #"&data_expira&"#) AND IN_Expira= TRUE) or IN_Expira= FALSE order by NO_Pasta Asc"
		RS_doc.Open SQL_doc, CON0

		if RS_doc.eof then%>
									<font class="style1">
										<%response.Write("<br><br><br><br><br>N&atilde;o existem documentos cadastrados!")%>
										</font>
									<%ELSE%>
									<select name="pasta" class="borda" id="pasta">
										<%
				conta_registros=0
				WHILE NOT RS_doc.eof
				total_pastas = RST("TOTAL_PASTAS")
				cod_tp_doc = RS_doc("CO_Pasta_Doc")		
				nom_tp_doc = RS_doc("NO_Pasta")	
								
				total_pastas=total_pastas*1
				if total_pastas = 1 then
					selected="selected"
				else
					tipo_arquivo=tipo_arquivo*1
					cod_tp_doc=cod_tp_doc*1					
					if tipo_arquivo= cod_tp_doc then
						selected="selected"		
						nome_pasta=nom_tp_doc			
					else
						selected=""		
					end if	
				end if					
				%>
										<option value="<%response.Write(cod_tp_doc)%>" <%response.Write(selected)%>>
											<%response.Write(nom_tp_doc)%>
											</option>
										<%
				conta_registros=conta_registros+1	
				RS_doc.MOVENEXT
				WEND
				%>
										</select>
									<%END if%>
								</div></td>
							</tr>
						</table></td>
					</tr>
					<tr>
						<td><hr></td>
					</tr>
					<tr>
						<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td width="33%">&nbsp;</td>
								<td width="34%">&nbsp;</td>
								<td width="33%" align="center"><input name="button" type="submit" class="botao_prosseguir" id="button" value="Prosseguir"></td>
							</tr>
						</table></td>
					</tr>
				</table></td></tr>
					<tr>					
					<td valign="top">
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<%
	
		Set RS_doc = Server.CreateObject("ADODB.Recordset")
		SQL_doc = "SELECT * FROM TB_Documentos where TP_Doc= "&tipo_arquivo&query&"order by DA_Doc Desc"		
		RS_doc.Open SQL_doc, CON_wf
%>
						<tr class="<%response.write(cor)%>">
							<td colspan="2" valign="top">&nbsp;</td>
						</tr>
						<tr class="<%response.write(cor)%>">
							<td colspan="2" valign="top" class="tb_subtit"><%response.Write(nome_pasta)%></td>
						</tr>
						<%		
						if RS_doc.eof then
						%>
						<tr class="<%response.write(cor)%>">
							<td colspan="2" valign="top"><div align="center"><font class="form_dado_texto">
									<%response.Write("Não existem documentos para este segmento!")%>
									</font></div></td>
						</tr>
						<%else
							check=2
							ordem=0
						%>
						<tr class="tb_fundo_linha_par">
						<%
							
							while not RS_doc.eof
							
							 if check mod 2 =0 then
							  cor = "tb_fundo_linha_par" 
							 else cor ="tb_fundo_linha_impar"
							  end if
		
								ordem=ordem+1						
								tit1=RS_doc("TI1_Doc")
								nome_arq=RS_doc("NO_Doc")
								extensao_arq = Array()
								extensao_arq= split(nome_arq, "." )
								extensao= extensao_arq(ubound(extensao_arq))
								nome_sessao="arq_"&ordem
								session(nome_sessao)=nome_arq
								
							
							
								select case extensao
								
									case "doc"
									icone="word"
									
									case "docx"
									icone="word"
									
									case "xls"
									icone="excel"
									
									case "xlsx"
									icone="excel"
									
									case "pdf"
									icone="pdf"
									
									case "pps"
									icone="pps"
									
									case "wmv"
									icone="wmv"
									
									case "wav"
									icone="wmv"
									
									case "avi"
									icone="avi"
									
									case "mpg"
									icone="mpg"
									
									case "mp3"
									icone="mpg"
									
									case "mpeg"
									icone="mpg"
									
									case "jpg"
									icone="jpg"
									
									case "jpeg"
									icone="jpg"
									
									case "gif"
									icone="gif"
									
									case "bmp"
									icone="bmp"
									
									case "rar"
									icone="zip"
									
									case "zip"
									icone="zip"
									
									case "txt"
									icone="word"
								
								end select
										
								data_de=RS_doc("DA_Doc")
								if data_de="" or isnull(data_de) then
								
								else			
									dados_dtd= split(data_de, "/" )
									dia_de= dados_dtd(0)
									mes_de= dados_dtd(1)
									ano_de= dados_dtd(2)
								end if
								dia_de=dia_de*1
								mes_de=mes_de*1								
								
								if dia_de<10 then
									dia_de="0"&dia_de
								end if
								if mes_de<10 then
									mes_de="0"&mes_de
								end if
								
								data_inicio=dia_de&"/"&mes_de&"/"&ano_de
						
						%>
							<td width="333" valign="top">
							<table width="100%" border="0" cellspacing="0" cellpadding="0">
									<tr valign="top" class="<%response.write(cor)%>">
										<td width="20"><img src="../../../../img/icones/<%=icone%>.gif" width="18" height="18"></td>
										<td><!--<a href="http://<%response.Write(site_escola)%>sndocs/download.asp?opt=<%response.Write(ordem)%>&ta=<%response.Write(tipo_arquivo)%>&al=<%response.Write(ano_letivo)%>&na=<%response.Write(nome_arq)%>" class="menu_sublista">-->
<a href="../pub/sndocs/download.asp?opt=<%response.Write(ordem)%>&ta=<%response.Write(tipo_arquivo)%>&al=<%response.Write(ano_letivo)%>&na=<%response.Write(nome_arq)%>" class="menu_sublista">										
											<%response.Write(tit1)%>
											</a></td>
									</tr>
									<tr valign="top" class="<%response.write(cor)%>">
										<td width="20">&nbsp;</td>
										<td><%response.Write("Publicado em "&data_inicio)
										  %></td>
										</tr>
								</table>
							</td>						
							<%  if ordem mod 3 <> 0 then
			'Quando o total de registros for impar a coluna abaixo é incluída.
									if RS_doc.eof then
								%><td width="333" valign="top">&nbsp;;
											</td>
										</tr>									
									<%
									end if	 
								else%>
									</tr>
									<tr class="<%response.write(cor)%>">	
										<%																			 
								end if	
						RS_doc.movenext					
						wend
					end if
%>																		
							
				</table>
			</td>
			
			</tr>
			
		</table></form>
	</td>
	</tr>
	<tr>
		<td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
	</tr>
	</table>

</body>
</html>
<%
if opt="a" then
			call GravaLog (chave,outro)
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