<!--#include file="dsn.asp"-->

<%
' Author Philippe Collignon
' Email PhCollignon@email.com


Response.Expires=0
Response.Buffer = TRUE
Response.Clear
'Response.BinaryWrite(Request.BinaryRead(Request.TotalBytes))
byteCount = Request.TotalBytes
'Response.BinaryWrite(Request.BinaryRead(varByteCount))

RequestBin = Request.BinaryRead(byteCount)
'a variável RequestBin recebe os valores do formulário
Dim UploadRequest
Set UploadRequest = CreateObject("Scripting.Dictionary")
'um dicionário de dados é criado para os valores do formulário

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
'esse é um objeto para manipulação de arquivos

Dim Conexao6,rs6,strSQL6

Set Conexao6 = Server.CreateObject("ADODB.Connection")

 Set rs6 = Server.CreateObject("ADODB.RecordSet")
 
 strSQL6 = "select * from contador where Cod_hits = 1" 


Conexao6.Open dsn
rs6.open strSQL6,Conexao6

dim nome

if rs6("hits") = "" or rs6("hits") = 0 then
nome = 1
end if


nome = int(rs6("hits")) + 1



BuildUploadRequest  RequestBin
'BuildUploadRequest é a função de upload.asp
'que pega os valores do formulário e divide-os
'dentro de um dicionário de dados.












 vProprietario = UploadRequest.Item("txt_proprietario").Item("Value")
  vEmail = UploadRequest.Item("txt_email").Item("Value")
 
  
  
   vTelefone = UploadRequest.Item("txt_telefone").Item("Value")
   
   
   
 
  
      vEndereco = UploadRequest.Item("txt_endereco").Item("Value")
	    
	 
	 
	 
	  
	  
     
	 
	  
	  
	  
	   
	   vCidade = UploadRequest.Item("combo1").Item("Value")
	      
	   	                                                            
	 
      vBairro = UploadRequest.Item("combo2").Item("Value")
	  	  
	  	  
      vTipo = UploadRequest.Item("txt_tipo").Item("Value")
	  	  	  	  
	  	 
	  vArea_Total = UploadRequest.Item("txt_a_total").Item("Value")
	    	  	  	  
	   						
		vArea_Construida = UploadRequest.Item("txt_a_constr").Item("Value")
				
												
		
		vQuartos = UploadRequest.Item("txt_quartos").Item("Value")
		
		
		 if vQuartos = "não informado" then
 vQuartos = "0"
 end if
		
		
      
	  vBanheiros = UploadRequest.Item("txt_banheiros").Item("Value")
	  
	  
      
	  vVagas = UploadRequest.Item("txt_vagas").Item("Value")
	  
	  
	  if vVagas = "não informado" then
 vVagas = "0"
 end if
	  
	  
	  
	  
	  
	  vNegociacao = UploadRequest.Item("txt_negociacao").Item("Value")
	  
	  dim vOcupacao
	  vOcupacao = UploadRequest.Item("txt_ocupacao").Item("Value")
	   						
									
		vValor = UploadRequest.Item("txt_valor").Item("Value")
		
		 vObs_imovel = UploadRequest.Item("obs_imovel").Item("Value")
		if vObs_imovel = "" then
		vObs_imovel = "sem observações"
		end if
		
		dim vObs_proprietario
		
		if vObs_proprietario = "" then
		vObs_proprietario = "sem observações"
		end if
		
		
		
        if vObs_proprietario = "" then
		vObs_proprietario = "sem observações"
		end if





if vEmail = "" then
  vEmail = "não informado"
  end if

if vTelefone = "" then
   vTelefone = "não informado"
   end if
 
if vArea_Total = "" then
   vArea_Total = "00"
   end if
   
   if vArea_Construida = "" then
   vArea_Construida = "00"
   end if

 'Acima é feito uma espécie de request dos formulários 
 'que consequentemente tem os valores colocados dentro 
 'das variáveis



























filepathname = UploadRequest.Item("blob").Item("FileName")
'a variável filepathname recebe o valor do endereço
'do arquivo mandado para upload

filename = Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
'aqui o nome do arquivo é separado do endereço completo.
filename = nome&filename





nome = int(nome) +1






If objFSO.FileExists(Server.MapPath(filename)) = True  then
response.redirect "form_incluir_imovel_internet03.asp"
else

'aqui é verificado se já existe algum arquivo 
'com o mesmo nome do arquivo uploadeado.
'se já existe então você é redirecionado para a
'página primeira_existente.html
'senão

value = UploadRequest.Item("blob").Item("Value")
'a variável value recebe o conteúdo do arquivo.

'Create FileSytemObject Component
 Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
'um objeto de manipulação de arquivos é instanciado

'Create and Write to a File
 pathEnd = Len(Server.mappath(Request.ServerVariables("PATH_INFO")))-14
 Set MyFile = ScriptObject.CreateTextFile(Left(Server.mappath(Request.ServerVariables("PATH_INFO")),pathEnd)&filename)
 'os valores do arquivo uploadeado são colocados numa variável para
 'manipulação de arquivos.
 
 For i = 1 to LenB(value)
	 MyFile.Write chr(AscB(MidB(value,i,1)))
 Next
 'o arquivo é colocado dentro de um arquivo de texto
 
 MyFile.Close
 
 
 '------------------------------segunda imagem --------------------------------
 
 
 
 '
 'a mesma coisa foi feita acima, os valores do arquivo uploadeado
 'são colocados num arquivo de texto.
 

Dim Conexao,vProprietario,vEndereco,vFoto_Grande,vFoto_Pequena,vLink_Foto,vCidade,vBairro,vTipo,vArea_Total,vArea_Construida,vQuartos,vBanheiros,vVagas,vNegociacao,vValor,vblob,vblob2,numBlob,numBlob2,vdata
 Dim varSucess_imovel,vEmail,vTelefone,vPresenca_primeira
 Dim vTexto_anuncio,vTitulo_anuncio
  
   Dim vdata2

vdata2 = now()

if len(vdata2) = 17 then
 vdata = left(now(),9)
 end if
 
 if len(vdata2) = 18 then
 vdata = left(now(),10)
 end if
 
 if len(vdata2) = 19 then
 vdata = left(now(),11)
 end if
   
 
									  
	
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	'um objeto conexão é instanciado
	
	Conexao.Open dsn
	'a conexão é aberta
	
	
	if vCidade <> "cqualquer" then
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 'um objeto recordset é instanciado
 
 SQL2 = "select * from combo1 where id_combo1 ="& vCidade
 
 rs2.open SQL2,Conexao,2,1
 'o recordset é aberto.
 dim vCidade2
 vCidade2 = rs2("nome_combo1")
 else
 vCidade2 = "não informado"
 end if
 
 
 
 if vBairro <> "bqualquer" then
 
 dim rs3,SQL3
 Set rs3 = Server.CreateObject("ADODB.RecordSet")
 SQL3 = "select * from combo2 where id_combo2 ="& vBairro
 
 rs3.open SQL3,Conexao,2,1
 dim vBairro2
 vBairro2 = rs3("nome_combo2")
 else
 vBairro2 = "não informado"
 end if
 
 
 
 
 
 
 
 
 
 	
 
	dim vVila_vend2
	dim vVila_vend
	 vVila_vend2=UploadRequest.Item("combo5").Item("Value")
	 session("vVila_vend2") = vVila_vend2
	 if session("vVila_vend2") = "" then
session("vVila_vend2") = request.querystring("vVila_vend2")

end if
	 
	 if session("vVila_vend2") <> "vlqualquer" then
	  dim rs88,SQL88
 Set rs88 = Server.CreateObject("ADODB.RecordSet")
 SQL88 = "select * from combo3 where id_combo3 ="& session("vVila_vend2")
 
 rs88.open SQL88,Conexao,2,1

 vVila_vend = rs88("nome_combo3")
 else
 vVila_vend = "não informado"
	end if                                      
									
	 	
	
	
	 dim vVila_comp2
	 dim vVila_comp
	 vVila_comp2=UploadRequest.Item("combo7").Item("Value")
	 session("vVila_comp2") = vVila_comp2
	 if session("vVila_comp2") = "" then
session("vVila_comp2") = request.querystring("vVila_comp2")

end if
	 
	 if session("vVila_comp2") <> "vlqualquer" then
	  dim rs99,SQL99
 Set rs99 = Server.CreateObject("ADODB.RecordSet")
 SQL99 = "select * from combo3 where id_combo3 ="&session("vVila_comp2")
 
 rs99.open SQL99,Conexao,2,1

 vVila_comp = rs99("nome_combo3")
 else
 vVila_comp = "não informado"
	end if                                      
	
	
 
 
 
 
 
 
 
 
  dim filename3,filename4,filename5
  filename3 = "imovel00000.jpg"
	 filename4 = "imovel00000.jpg"
  filename5 = "imovel00000.jpg"
	 filename6 = "imovel00000.jpg"
	  filename7 = "imovel00000.jpg"
	 
	 
	Conexao.execute"Insert into imoveis(proprietario,endereco,telefone,email,foto_grande,foto_pequena,foto_grande1,foto_grande2,foto_grande3,foto_grande4,foto_grande5,link_foto,cidade,bairro,tipo,area_total,area_construida,quartos,banheiros,vagas,negociacao,valor,data,obs_imovel,obs_proprietario,presenca_primeira,titulo_anuncio,texto_anuncio,StandBy,ocupacao,vila,captacao,data_atualizacao) values( '"& vProprietario &"','"& vEndereco &"','"& vTelefone &"','"& vEmail &"','"& filename &"','"& filename &"','"& filename &"','"& filename3 &"','"& filename4 &"','"& filename5 &"','"& filename7 &"','"& "icon_foto.gif" &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vArea_Total &"','"& vArea_Construida &"','"& vQuartos &"','"& vBanheiros &"','"& vVagas &"','"& vNegociacao &"','"& vValor &"','"& now() &"','"& vObs_imovel &"','"& vObs_proprietario &"','"& "excluido" &"','"& "não informado" &"','"& "não informado" &"','"& "excluido" &"','"& vOcupacao &"','"& vVila_vend &"','"& "internet" &"','"& now() &"')"
	'aqui os dados são colocados dentro do banco de dados.
	dim rs5,SQL5
 Set rs5 = Server.CreateObject("ADODB.RecordSet")
 SQL5 = "select * from imoveis ORDER BY cod_imovel DESC"
 rs5.open SQL5,Conexao,2,1
	 rs5.movefirst
	 dim varCodImovel5 
	 varCodImovel5 = rs5("cod_imovel")
	'aqui a variável varCodImovel5 recebe o valor do código do imóvel
	
	
	Conexao6.execute"update contador set hits='"&nome&"' where Cod_hits = 1"
	 
	
	
	 dim vPergunta
	 
	 vPergunta = UploadRequest.Item("txt_pergunta").Item("Value")
	 if vPergunta = "sim" then
	  
	  dim vVagas_comp
	   vVagas_comp = UploadRequest.Item("txt_vagas_comp").Item("Value")
	  
	  
	  if vVagas_comp = "não informado" then
 vVagas_comp = "0"
 end if
	  
	  
	  
	 
	 dim vCidade_comp,vBairro_comp
	 dim vCidade2_comp,vBairro2_comp
	 dim vTipo_comp,vQuartos_comp,vValor_comp
	 dim vDescricao_comp
	 
	 vCidade_comp=UploadRequest.Item("combo3").Item("Value")
	 vBairro_comp=UploadRequest.Item("combo4").Item("Value")
	 vTipo_comp=UploadRequest.Item("txt_tipo_comp").Item("Value")
	 
	 
	 
	 vQuartos_comp=UploadRequest.Item("txt_quartos_comp").Item("Value")
	 
	 
	  if vQuartos_comp = "não informado" then
 vQuartos_comp = "0"
 end if
	 
	 
	 
	 
	 
	 
	 vValor_comp=UploadRequest.Item("txt_valor_comp").Item("Value")
	 vDescricao_comp=UploadRequest.Item("txt_descricao_comp").Item("Value")
	 
	 if vValor_comp="" then
	 vValor_comp="00"
	 end if
	 
	 if vDescricao_comp="" then
	 vDescricao_comp="não informado"
	 end if
	 
	 if vCidade_comp <> "cqualquer" then
	 dim rs55,SQL55
 Set rs55 = Server.CreateObject("ADODB.RecordSet")
 'um objeto recordset é instanciado
 
 SQL55 = "select * from combo1 where id_combo1 ="& vCidade_comp
 
 rs55.open SQL55,Conexao,2,1
 'o recordset é aberto.
 
 vCidade2_comp = rs55("nome_combo1")
 else
 vCidade2_comp = "não informado"
 end if
 
 
 
 if vBairro_comp <> "bqualquer" then
 dim rs44,SQL44
 Set rs44 = Server.CreateObject("ADODB.RecordSet")
 SQL44 = "select * from combo2 where id_combo2 ="& vBairro_comp
 
 rs44.open SQL44,Conexao,2,1


 vBairro2_comp = rs44("nome_combo2")
 
 else
 vBairro2_comp="não informado"
 end if
 
 
 
 
 dim vimagem,varCodImovel,vLink
 vimagem = "imovel00000.jpg"
 varCodImovel = "00"
 vLink= "não informado"
 Conexao.execute"Insert into permuta(Foto_imovel, Nome, Email, Telefone,endereco_vend, cidade_vend, bairro_vend,tipo_vend,descricao_vend,cidade_comp,bairro_comp,tipo_comp,descricao_comp,cod_imovel,link_imovel,data,quartos_vend,quartos_comp,valor_vend,valor_comp,vila_vend,vila_comp,atendimento,data_atualizacao,vagas_vend,vagas_comp) values( '"& filename &"','"& vProprietario &"','"& vEmail &"','"& vTelefone &"','"& vEndereco &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vObs_imovel &"','"& vCidade2_comp &"','"& vBairro2_comp &"','"& vTipo_comp &"','"& vDescricao_comp &"','"& varCodImovel &"','"& vLink &"','"& now() &"','"& vQuartos &"','"& vQuartos_comp &"','"& vValor &"','"& vValor_comp &"','"& vVila_vend &"','"& vVila_comp &"','"& "internet" &"','"& now() &"','"& vVagas &"','"& vVagas_comp &"')"
 
  Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,atendimento,data_atualizacao,vila,vagas,ocupacao) values( '"& vProprietario &"','"& vTelefone &"','"& vEmail &"','"& vCidade2_comp &"','"& vBairro2_comp &"','"& vTipo_comp &"','"& vQuartos_comp &"','"& "Compra" &"','"& vValor_comp &"','"& now() &"','"& vDescricao_comp &"','"& "internet" &"','"& now() &"','"& vVila_comp &"','"& vVagas_comp &"','"& "não informado" &"')"
 
 
			
 dim rs555,SQL555
 Set rs555 = Server.CreateObject("ADODB.RecordSet")
 SQL555 = "select * from permuta ORDER BY cod_permuta ASC" 
	
	rs555.open SQL555,Conexao,2,1  
	 rs555.moveLast
	
	dim varCodPermuta555
	
	
	varCodPermuta555 = rs555("cod_permuta")
	
	
	 dim rs333,SQL333
 Set rs333 = Server.CreateObject("ADODB.RecordSet")
 SQL333 = "select * from compradores ORDER BY Cod_compradores ASC" 
	
	rs333.open SQL333,Conexao,2,1  
	 rs333.moveLast
	dim varNome
	dim varTelefone
	dim varCodComprador
	
	varNome = rs333("nome")
	varTelefone = rs333("telefone")
	varCodComprador = rs333("cod_compradores")
	 
 
 
 
	end if
	
			
 dim rs444,SQL444
 Set rs444 = Server.CreateObject("ADODB.RecordSet")
 SQL444 = "select * from imoveis ORDER BY cod_imovel ASC" 
	
	rs444.open SQL444,Conexao,2,1  
	 rs444.moveLast
	
	dim varCodImovel444
	
	
	varCodImovel444 = rs444("cod_imovel")	
	
	

	 
	 
	  
	 
	 
	 
	 response.Redirect "mostrar_conta02.asp?varNome="&vProprietario&"&varTelefone="&vTelefone&"&varCodComprador="&varCodComprador&"&vPergunta="&vPergunta&"&varCodImovel="&varCodImovel444&"&varCodPermuta="&varCodPermuta555&""






 
 
 
 
 
 
 
 
 
 
 
 
 
 
%>
<b>Uploaded file : </b><%=filename%><BR>
<img src="<%=filename%>"><br>

<!--#include file="upload.asp"-->
 
 <%
 end if
  
      
 
 Conexao6.Close
		   rs6.close
           
           Conexao.Close
		   rs2.close
		   rs3.close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>

