<!--#include file="dsn.asp"-->
<!--#include file="loggedin.asp"-->
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



vTitulo_anuncio = UploadRequest.Item("txt_titulo").Item("Value")

if vTitulo_anuncio = "" then
vTitulo_anuncio = "não informado"
end if 


vTexto_anuncio = UploadRequest.Item("txt_anuncio").Item("Value")
 if vTexto_anuncio = "" then
 vTexto_anuncio = "não informado"
 end if




 vProprietario = UploadRequest.Item("txt_proprietario").Item("Value")
  vEmail = UploadRequest.Item("txt_email").Item("Value")
  vPresenca_primeira = UploadRequest.Item("txt_presenca_primeira").Item("Value")
  
  
   vTelefone = UploadRequest.Item("txt_telefone").Item("Value")
   
   
   
 
  
      vEndereco = UploadRequest.Item("txt_endereco").Item("Value")
	    
	 
	 
	 
	  
	  
     
	  vLink_Foto = UploadRequest.Item("txt_link_foto").Item("Value")
	  
	  
	  
	   
	   vCidade = UploadRequest.Item("combo1").Item("Value")
	      
	   	                                                            
	 
      vBairro = UploadRequest.Item("combo2").Item("Value")
	  
	  vVila =  UploadRequest.Item("combo5").Item("Value")
	  	  
	  	  
      vTipo = UploadRequest.Item("txt_tipo").Item("Value")
	  	  	  	  
	  	 
	  vArea_Total = UploadRequest.Item("txt_a_total").Item("Value")
	    	  	  	  
	   						
		vArea_Construida = UploadRequest.Item("txt_a_constr").Item("Value")
				
												
		
		vQuartos = UploadRequest.Item("txt_quartos").Item("Value")
		
		 if vQuartos = "não informado" then
 vQuartos = "0"
 end if
		
      
	  vBanheiros = UploadRequest.Item("txt_banheiros").Item("Value")
	  
	  
      
	  vVagas = UploadRequest.Item("txt_vagas").Item("Value")
	  
	  if vVagas = "não informado" or vVagas = "qualquer um" or vVagas = "" then
	  vVagas = "0"
	  end if
	  
	  
	  
	  vNegociacao = UploadRequest.Item("txt_negociacao").Item("Value")
	   						
									
		vValor = UploadRequest.Item("txt_valor").Item("Value")
		
		dim vObs_imovel
        
		 vObs_imovel = UploadRequest.Item("obs_imovel").Item("Value")
		
		if vObs_imovel = "" then
		vObs_imovel = "sem observações"
		end if
		
		dim vObs_proprietario
		
		 vObs_proprietario = UploadRequest.Item("obs_proprietario").Item("Value")
		
		
		
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

dim StandBy,vOcupacao

 StandBy = UploadRequest.Item("txt_standby").Item("Value")
 vOcupacao = UploadRequest.Item("txt_ocupacao").Item("Value")
 
 dim vQualidade
 
 vQualidade = UploadRequest.Item("txt_qualidade").Item("Value") 
 
 
 dim vCaptacao
 
 vCaptacao = UploadRequest.Item("txt_captacao").Item("Value")
 
 if vCaptacao = "" then
 vCaptacao = "não informado"
  end if
  

























filepathname = UploadRequest.Item("blob").Item("FileName")
'a variável filepathname recebe o valor do endereço
'do arquivo mandado para upload

filename = Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
'aqui o nome do arquivo é separado do endereço completo.
filename = nome&filename

filepathname2 = UploadRequest.Item("blob2").Item("FileName")



nome = int(nome) +1

filename2 = Right(filepathname2,Len(filepathname2)-InstrRev(filepathname2,"\"))
'
'a mesma coisa é feita aqui com relação aos
'arquivos mandados via upload
filename2 = nome&filename2




If objFSO.FileExists(Server.MapPath(filename)) = True or objFSO.FileExists(Server.MapPath(filename2)) = True then
response.redirect "primeira_existente.html"
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
 
 
 contentType = UploadRequest.Item("blob2").Item("ContentType")

value = UploadRequest.Item("blob2").Item("Value")

'Create FileSytemObject Component
 Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")

'Create and Write to a File
 pathEnd = Len(Server.mappath(Request.ServerVariables("PATH_INFO")))-14
 Set MyFile2 = ScriptObject.CreateTextFile(Left(Server.mappath(Request.ServerVariables("PATH_INFO")),pathEnd)&filename2)
 
 For i = 1 to LenB(value)
	 MyFile2.Write chr(AscB(MidB(value,i,1)))
 Next
 
 MyFile2.Close
 
 '
 'a mesma coisa foi feita acima, os valores do arquivo uploadeado
 'são colocados num arquivo de texto.
 

Dim Conexao,vProprietario,vEndereco,vFoto_Grande,vFoto_Pequena,vLink_Foto,vCidade,vBairro,vTipo,vArea_Total,vArea_Construida,vQuartos,vBanheiros,vVagas,vNegociacao,vValor,vblob,vblob2,numBlob,numBlob2,vdata
 Dim varSucess_imovel,vEmail,vTelefone,vPresenca_primeira
 Dim vTexto_anuncio,vTitulo_anuncio
  
   Dim vdata2

vdata = now()
vdata2 = now()

   
 
									  
	
	
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
 
 
 if vVila <> "vlqualquer" then
 dim rs333,SQL333
 Set rs333 = Server.CreateObject("ADODB.RecordSet")
 SQL333 = "select * from combo3 where id_combo3 ="&vVila
 
 rs333.open SQL333,Conexao,2,1
 dim vVila2
 vVila2 = rs333("nome_combo3")
 
 else
 
 vVila2 = "não informado"
 
 end if
 
 
 
 
 
 
  dim filename3,filename4,filename5
  filename3 = "imovel00000.jpg"
	 filename4 = "imovel00000.jpg"
  filename5 = "imovel00000.jpg"
	 filename6 = "imovel00000.jpg"
	  filename7 = "imovel00000.jpg"
	 
	 
	Conexao.execute"Insert into imoveis(proprietario,endereco,telefone,email,foto_grande,foto_pequena,foto_grande1,foto_grande2,foto_grande3,foto_grande4,foto_grande5,link_foto,cidade,bairro,tipo,area_total,area_construida,quartos,banheiros,vagas,negociacao,valor,data,obs_imovel,obs_proprietario,presenca_primeira,titulo_anuncio,texto_anuncio,StandBy,ocupacao,captacao,data_atualizacao,vila,qualidade) values( '"& vProprietario &"','"& vEndereco &"','"& vTelefone &"','"& vEmail &"','"& filename &"','"& filename2 &"','"& filename &"','"& filename4 &"','"& filename5 &"','"& filename6 &"','"& filename7 &"','"& vLink_Foto &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vArea_Total &"','"& vArea_Construida &"','"& vQuartos &"','"& vBanheiros &"','"& vVagas &"','"& vNegociacao &"','"& vValor &"','"& vdata &"','"& vObs_imovel &"','"& vObs_proprietario &"','"& vPresenca_primeira &"','"& vTitulo_anuncio &"','"& vTexto_anuncio &"','"& StandBy &"','"& vOcupacao &"','"& vCaptacao &"','"& vData2 &"','"& vVila2 &"','"& vQualidade &"')"
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
	 
	
	  
      response.Redirect "form_incluir_imovel.asp?varSucesso_imovel="&vProprietario&""
      'você é redirecionado para o formulário de inclusão de imóvel.







 
 
 
 
 
 
 
 
 
 
 
 
 
 
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

