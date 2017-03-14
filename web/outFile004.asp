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









Dim vCidade_vend,vBairro_vend,vTipo_vend,vDescricao_vend
 Dim vCidade_comp,vBairro_comp,vTipo_comp,vDescricao_comp,vLink
 Dim vProprietario,vTelefone,vEmail,vEndereco
 dim vQuartos_vend,vQuartos_comp,vValor_vend,vValor_comp
 dim vVagas_vend,vVagas_comp
 dim vPergunta
 
 
 
 
 vProprietario=UploadRequest.Item("txt_proprietario").Item("Value")
 vTelefone=UploadRequest.Item("txt_telefone").Item("Value")
 vEmail=UploadRequest.Item("txt_email").Item("Value")
 vEndereco=UploadRequest.Item("txt_endereco").Item("Value")
 vCidade_vend=UploadRequest.Item("combo1").Item("Value")
 vBairro_vend=UploadRequest.Item("combo2").Item("Value")
 vTipo_vend=UploadRequest.Item("txt_tipo_vend").Item("Value")
 vDescricao_vend=UploadRequest.Item("txt_descricao_vend").Item("Value")
vVagas_vend=UploadRequest.Item("txt_vagas_vend").Item("Value")
 vVagas_comp=UploadRequest.Item("txt_vagas_comp").Item("Value")
 
 
 vCidade_comp=UploadRequest.Item("combo3").Item("Value")
 vBairro_comp=UploadRequest.Item("combo4").Item("Value")
 vTipo_comp=UploadRequest.Item("txt_tipo_comp").Item("Value")
 vDescricao_comp=UploadRequest.Item("txt_descricao_comp").Item("Value")
 vQuartos_comp=UploadRequest.Item("txt_quartos_comp").Item("Value")
 vQuartos_vend=UploadRequest.Item("txt_quartos_vend").Item("Value")
 
 vValor_vend=UploadRequest.Item("txt_valor_vend").Item("Value")
 vValor_comp=UploadRequest.Item("txt_valor_comp").Item("Value")
 
 
 vLink="não informado"
 
 
 
 varCodimovel = "00"
   vimagem = "imovel00000.jpg"
  
  
  dim vdata2
  vdata2 = now()

if len(vdata2) = 17 then
 vdata = left(now(),9)
 end if

if vEmail = "" then
vEmail = "não informado"
end if

  
 if vVagas_vend = "não informado" then
 vVagas_vend = "0"
 end if
 
 if vVagas_comp = "não informado" then
 vVagas_comp = "0"
 end if
 
 if vQuartos_vend = "não informado" then
 vQuartos_vend = "0"
 end if
 
 if vQuartos_comp = "não informado" then
 vQuartos_comp = "0"
 end if
 
  
 
	  
	  if vEmail = "" then
	  vEmail = "não informado"
	  end if
	  
	  													  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	
	if vCidade_vend <> "cqualquer" then
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 'um objeto recordset é instanciado
 
 SQL2 = "select * from combo1 where id_combo1 ="& vCidade_vend
 
 rs2.open SQL2,Conexao,2,1
 'o recordset é aberto.
 dim vCidade2_vend
 vCidade2_vend = rs2("nome_combo1")
 else
 vCidade2_vend="não informado"
 end if
 
 
 
 if vBairro_vend <> "bqualquer" then
 
 dim rs3,SQL3
 Set rs3 = Server.CreateObject("ADODB.RecordSet")
 SQL3 = "select * from combo2 where id_combo2 ="& vBairro_vend
 
 rs3.open SQL3,Conexao,2,1
 dim vBairro2_vend
 vBairro2_vend = rs3("nome_combo2")
 else
 vBairro2_vend = "não informado"
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
 SQL99 = "select * from combo3 where id_combo3 ="& session("vVila_comp2")
 
 rs99.open SQL99,Conexao,2,1

 vVila_comp = rs99("nome_combo3")
 else
 vVila_comp = "não informado"
	end if                                      
	
	
	
	
	if vCidade_comp <> "cqualquer" then
	dim rs5,SQL5
 Set rs5 = Server.CreateObject("ADODB.RecordSet")
 'um objeto recordset é instanciado
 
 SQL5 = "select * from combo1 where id_combo1 ="& vCidade_comp
 
 rs5.open SQL5,Conexao,2,1
 'o recordset é aberto.
 dim vCidade2_comp
 vCidade2_comp = rs5("nome_combo1")
 else
 vCidade2_comp = "não informado"
 end if
 
 
 
 if vBairro_comp <> "bqualquer" then
 
 dim rs4,SQL4
 Set rs4 = Server.CreateObject("ADODB.RecordSet")
 SQL4 = "select * from combo2 where id_combo2 ="& vBairro_comp
 
 rs4.open SQL4,Conexao,2,1
 dim vBairro2_comp
 vBairro2_comp = rs4("nome_combo2")
 else
 vBairro2_comp = "não informado"
 end if
 
	
	
	
	
	
	
	
	
	dim varCodimovel2
	varCodimovel2 = "não informado"
	
	
 
 dim vFoto2,vFoto3,vFoto4,vFoto5,vFoto6,vFoto7
 
 vFoto = "imovel00000.jpg"
 vFoto2 = "imovel00000.jpg"
 vFoto3 = "imovel00000.jpg"
 vFoto4 = "imovel00000.jpg"
 vFoto5 = "imovel00000.jpg"
 vFoto6 = "imovel00000.jpg"
 vFoto7 = "imovel00000.jpg"
























filepathname = UploadRequest.Item("blob").Item("FileName")
'a variável filepathname recebe o valor do endereço
'do arquivo mandado para upload

filename = Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
'aqui o nome do arquivo é separado do endereço completo.
filename = nome&filename





nome = int(nome) +1






If objFSO.FileExists(Server.MapPath(filename)) = True  then
response.redirect "form_permuta_incluir04.asp"
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
 

	'Conexao.execute"Insert into permuta(Foto_imovel, Nome, Email, Telefone,endereco_vend, cidade_vend, bairro_vend,tipo_vend,descricao_vend,cidade_comp,bairro_comp,tipo_comp,descricao_comp,cod_imovel,link_imovel,data,quartos_vend,quartos_comp,valor_vend,valor_comp,vila_vend,vila_comp,vagas_vend,vagas_comp) values( '"& filename &"','"& vProprietario &"','"& vEmail &"','"& vTelefone &"','"& vEndereco &"','"& vCidade2_vend &"','"& vBairro2_vend &"','"& vTipo_vend &"','"& vDescricao_vend &"','"& vCidade2_comp &"','"& vBairro2_comp &"','"& vTipo_comp &"','"& vDescricao_comp &"','"& varCodImovel &"','"& vLink &"','"& vData2 &"','"& vQuartos_vend &"','"& vQuartos_comp &"','"& vValor_vend &"','"& vValor_comp &"','"& vVila_vend &"','"& vVila_comp &"','"& vVagas_vend &"','"& vVagas_comp &"')" 
	Conexao.execute"Insert into permuta(Foto_imovel, Nome, Email, Telefone,endereco_vend, cidade_vend, bairro_vend,tipo_vend,descricao_vend,cidade_comp,bairro_comp,tipo_comp,descricao_comp,cod_imovel,link_imovel,data,quartos_vend,quartos_comp,valor_vend,valor_comp,atendimento,data_atualizacao,vila_vend,vila_comp,vagas_vend,vagas_comp) values( '"& filename &"','"& vProprietario &"','"& vEmail &"','"& vTelefone &"','"& vEndereco &"','"& vCidade2_vend &"','"& vBairro2_vend &"','"& vTipo_vend &"','"& vDescricao_vend &"','"& vCidade2_comp &"','"& vBairro2_comp &"','"& vTipo_comp &"','"& vDescricao_comp &"','"& varCodImovel &"','"& vLink &"','"& now() &"','"& vQuartos_vend &"','"& vQuartos_comp &"','"& vValor_vend &"','"& vValor_comp &"','"& "internet" &"','"& now() &"','"& vVila_vend &"','"& vVila_comp &"','"& vVagas_vend &"','"& vVagas_comp &"')"
	
	
	
	'Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,vila,vagas,ocupacao) values( '"& vProprietario &"','"& vTelefone &"','"& vEmail &"','"& vCidade2_comp &"','"& vBairro2_comp &"','"& vTipo_comp &"','"& vQuartos_comp &"','"& "compra" &"','"& vValor_comp &"','"& vdata &"','"& vDescricao_comp &"','"& vVila_comp &"','"& vVagas_comp &"','"& "não informado" &"')"
	Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,atendimento,data_atualizacao,vila,vagas,ocupacao) values( '"& vProprietario &"','"& vTelefone &"','"& vEmail &"','"& vCidade2_comp &"','"& vBairro2_comp &"','"& vTipo_comp &"','"& vQuartos_comp &"','"& "compra" &"','"& vValor_comp &"','"& now() &"','"& vDescricao_comp &"','"& "internet" &"','"& now() &"','"& vVila_comp &"','"& vVagas_comp &"','"& "não informado" &"')"
	
	 
	 vPergunta = UploadRequest.Item("txt_pergunta").Item("Value")
	 
	 
	 if vPergunta = "sim" then
	 
	  
	
	 
	 Conexao.execute"Insert into imoveis(proprietario,endereco,telefone,email,foto_grande,foto_pequena,foto_grande1,foto_grande2,foto_grande3,foto_grande4,foto_grande5,link_foto,cidade,bairro,tipo,area_total,area_construida,quartos,banheiros,vagas,negociacao,valor,data,obs_imovel,obs_proprietario,presenca_primeira,titulo_anuncio,texto_anuncio,standby,ocupacao,vila,data_atualizacao) values( '"& vProprietario &"','"& vEndereco &"','"& vTelefone &"','"& vEmail &"','"& filename &"','"& filename &"','"& filename &"','"& vFoto4 &"','"& vFoto5 &"','"& vFoto6 &"','"& vFoto7 &"','"& "icon_foto.gif" &"','"& vCidade2_vend&"','"& vBairro2_vend &"','"& vTipo_vend &"','"& "00" &"','"& "00" &"','"& vQuartos_vend &"','"& "00" &"','"& vVagas_vend &"','"& "venda" &"','"& vValor_vend &"','"& now() &"','"& vDescricao_vend &"','"& "não informado" &"','"& "excluido" &"','"& "não informado" &"','"& "não informado" &"','"& "excluido" &"','"& "não informado" &"','"& vVila_vend &"','"& now() &"')"
	
	 		
 dim rs444,SQL444
 Set rs444 = Server.CreateObject("ADODB.RecordSet")
 SQL444 = "select * from imoveis ORDER BY cod_imovel ASC" 
	
	rs444.open SQL444,Conexao,2,1  
	 rs444.moveLast
	
	dim varCodImovel444
	
	
	varCodImovel444 = rs444("cod_imovel")	
	
	 
	
	
	
	
	 end if
	 
	
	 
     
	  
	Conexao6.execute"update contador set hits='"&nome&"' where Cod_hits = 1"
	 
					
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
	 
	
	
	 response.Redirect "mostrar_conta03.asp?varNome="&vProprietario&"&varTelefone="&vTelefone&"&varCodComprador="&varCodComprador&"&vPergunta="&vPergunta&"&varCodImovel="&varCodImovel444&"&varCodPermuta="&varCodPermuta555&""
     






 
 
 
 
 
 
 
 
 
 
 
 
 
 
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

