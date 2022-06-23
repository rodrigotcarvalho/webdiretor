<%
function pais(cod_cons,caminho)

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& caminho & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Paises WHERE CO_Pais ="& cod_cons
		RS.Open SQL, CON

			if RS.EOF then
				pais = "ERR1"
			else
				pais = RS("NO_Pais")
			end if

		RS.close			
		set RS = nothing
	  
		CON.close
		set CON = nothing		  
			
end function

function nacionalidade(cod_cons,caminho)

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& caminho & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Nacionalidades WHERE CO_Nacionalidade ="& cod_cons
		RS.Open SQL, CON

			if RS.EOF then
				nacionalidade = "ERR1"
			else
				nacionalidade = RS("TX_Nacionalidade")
			end if

		RS.close			
		set RS = nothing
	  
		CON.close
		set CON = nothing	
			
end function

function uf(cod_cons,caminho)

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& caminho & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT * FROM TB_UF WHERE SG_UF ='"& cod_cons&"'"
		RS.Open SQL, CON

			if RS.EOF then
				uf = "ERR1"
			else
				uf = RS("NO_UF")
			end if
			
		RS.close			
		set RS = nothing
	  
		CON.close
		set CON = nothing	
					
end function

function municipio(cod_cons,caminho,uf)

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& caminho & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL6 = "SELECT * FROM TB_Municipios WHERE SG_UF ='"& uf&"' AND CO_Municipio = "&cod_cons
		RS.Open SQL, CON

			if RS.EOF then
				municipio = "ERR1"
			else
				municipio = RS("NO_Municipio")
			end if
			
		RS.close			
		set RS = nothing
	  
		CON.close
		set CON = nothing				
			
end function

function bairro(cod_cons,caminho,uf,municipio)

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& caminho & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL6 = "SELECT * FROM TB_Bairros WHERE CO_Bairro ="& cod_cons &"AND SG_UF ='"& uf&"' AND CO_Municipio = "&municipio
		RS.Open SQL, CON

			if RS.EOF then
				bairro = "ERR1"
			else
				bairro = RS("NO_Bairro")
			end if
			
		RS.close			
		set RS = nothing
	  
		CON.close
		set CON = nothing				
			
end function

function ocupacao(cod_cons,caminho)

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& caminho & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Ocupacoes WHERE CO_Ocupacao ="& cod_cons
		R0.Open SQL, CON

			if RS.EOF then
				ocupacao = "ERR1"
			else
				ocupacao= RS("NO_Ocupacao")
			end if

		RS.close			
		set RS = nothing
	  
		CON.close
		set CON = nothing	

end function


function religiao(cod_cons,caminho)

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& caminho & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Religiao WHERE CO_Religiao ="& cod_cons
		R0.Open SQL, CON

			if RS.EOF then
				religiao = "ERR1"
			else
				religiao = RS0("TX_Descricao_Religiao")
			end if

		RS.close			
		set RS = nothing
	  
		CON.close
		set CON = nothing	

end function

function estado_civil(cod_cons,caminho)

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& caminho & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Estado_Civil WHERE CO_Estado_Civil ='"& cod_cons&"'"
		R0.Open SQL, CON

			if RS.EOF then
				estado_civil = "ERR1"
			else
				estado_civil= RS("TX_Estado_Civil")
			end if

		RS.close			
		set RS = nothing
	  
		CON.close
		set CON = nothing	

end function

function raca(cod_cons,caminho)

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& caminho & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Raca WHERE CO_Raca ="& raca
		RS.Open SQL, CON

		if RS.EOF then
			raca = "ERR1"
		else
			raca = RS1("TX_Descricao_Raca")
		end if

		RS.close			
		set RS = nothing
	  
		CON.close
		set CON = nothing	

end function