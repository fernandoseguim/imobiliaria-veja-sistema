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
'a vari�vel RequestBin recebe os valores do formul�rio
Dim UploadRequest
Set UploadRequest = CreateObject("Scripting.Dictionary")
'um dicion�rio de dados � criado para os valores do formul�rio

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
'esse � um objeto para manipula��o de arquivos



 
 







BuildUploadRequest  RequestBin
'BuildUploadRequest � a fun��o de upload.asp
'que pega os valores do formul�rio e divide-os
'dentro de um dicion�rio de dados.



vTitulo_anuncio = UploadRequest.Item("txt_titulo").Item("Value")

if vTitulo_anuncio = "" then
vTitulo_anuncio = "n�o informado"
end if 


vTexto_anuncio = UploadRequest.Item("txt_anuncio").Item("Value")
 if vTexto_anuncio = "" then
 vTexto_anuncio = "n�o informado"
 end if




 vProprietario = UploadRequest.Item("txt_proprietario").Item("Value")
  vEmail = UploadRequest.Item("txt_email").Item("Value")
  vPresenca_primeira = UploadRequest.Item("txt_presenca_primeira").Item("Value")
  
  
   vTelefone = UploadRequest.Item("txt_telefone").Item("Value")
   
   
   
 
  
      vEndereco = UploadRequest.Item("txt_endereco").Item("Value")
	    
	 
	 
	 
	  
	  
     
	  vLink_Foto = UploadRequest.Item("txt_link_foto").Item("Value")
	  
	  
	  
	   
	   vCidade = UploadRequest.Item("combo1").Item("Value")
	      
	   	                                                            
	 
      vBairro = UploadRequest.Item("combo2").Item("Value")
	  	  
	  	  
      vTipo = UploadRequest.Item("txt_tipo").Item("Value")
	  	  	  	  
	  	 
	  vArea_Total = UploadRequest.Item("txt_a_total").Item("Value")
	    	  	  	  
	   						
		vArea_Construida = UploadRequest.Item("txt_a_constr").Item("Value")
				
												
		
		vQuartos = UploadRequest.Item("txt_quartos").Item("Value")
		
		
		
      
	  vBanheiros = UploadRequest.Item("txt_banheiros").Item("Value")
	  
	  
      
	  vVagas = UploadRequest.Item("txt_vagas").Item("Value")
	  
	  
	  
	  
	  vNegociacao = UploadRequest.Item("txt_negociacao").Item("Value")
	   						
									
		vValor = UploadRequest.Item("txt_valor").Item("Value")
		
		
		dim vOcupacao,standBy
		
		vOcupacao = UploadRequest.Item("txt_ocupacao").Item("Value")
		
		standBy = UploadRequest.Item("txt_standby").Item("Value")
		
		dim vObs_imovel
        
		 vObs_imovel = UploadRequest.Item("obs_imovel").Item("Value")
		
		if vObs_imovel = "" then
		vObs_imovel = "sem observa��es"
		end if
		
		dim vObs_proprietario
		
		 vObs_proprietario = UploadRequest.Item("obs_proprietario").Item("Value")
		
		
		
        if vObs_proprietario = "" then
		vObs_proprietario = "sem observa��es"
		end if





if vEmail = "" then
  vEmail = "n�o informado"
  end if

if vTelefone = "" then
   vTelefone = "n�o informado"
   end if
 
if vArea_Total = "" then
   vArea_Total = "n�o informado"
   end if
   
   if vArea_Construida = "" then
   vArea_Construida = "n�o informado"
   end if

 'Acima � feito uma esp�cie de request dos formul�rios 
 'que consequentemente tem os valores colocados dentro 
 'das vari�veis


 
 

























filepathname = UploadRequest.Item("blob").Item("FileName")
'a vari�vel filepathname recebe o valor do endere�o
'do arquivo mandado para upload

filename = Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
'aqui o nome do arquivo � separado do endere�o completo.

filepathname2 = UploadRequest.Item("blob2").Item("FileName")
filename2 = Right(filepathname2,Len(filepathname2)-InstrRev(filepathname2,"\"))
'a mesma coisa � feita aqui com rela��o aos
'arquivos mandados via upload

If objFSO.FileExists(Server.MapPath(filename)) = True or objFSO.FileExists(Server.MapPath(filename2)) = True then
response.redirect "primeira_existente.html"
else

'aqui � verificado se j� existe algum arquivo 
'com o mesmo nome do arquivo uploadeado.
'se j� existe ent�o voc� � redirecionado para a
'p�gina primeira_existente.html
'sen�o

value = UploadRequest.Item("blob").Item("Value")
'a vari�vel value recebe o conte�do do arquivo.

'Create FileSytemObject Component
 Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
'um objeto de manipula��o de arquivos � instanciado

'Create and Write to a File
 pathEnd = Len(Server.mappath(Request.ServerVariables("PATH_INFO")))-14
 Set MyFile = ScriptObject.CreateTextFile(Left(Server.mappath(Request.ServerVariables("PATH_INFO")),pathEnd)&filename)
 'os valores do arquivo uploadeado s�o colocados numa vari�vel para
 'manipula��o de arquivos.
 
 For i = 1 to LenB(value)
	 MyFile.Write chr(AscB(MidB(value,i,1)))
 Next
 'o arquivo � colocado dentro de um arquivo de texto
 
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
 's�o colocados num arquivo de texto.
 

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
	'um objeto conex�o � instanciado
	
	Conexao.Open dsn
	'a conex�o � aberta
	
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 'um objeto recordset � instanciado
 
 SQL2 = "select * from combo1 where id_combo1 ="& vCidade
 
 rs2.open SQL2,Conexao,2,1
 'o recordset � aberto.
 dim vCidade2
 vCidade2 = rs2("nome_combo1")
 
 
 dim rs3,SQL3
 Set rs3 = Server.CreateObject("ADODB.RecordSet")
 SQL3 = "select * from combo2 where id_combo2 ="& vBairro
 
 rs3.open SQL3,Conexao,2,1
 dim vBairro2
 vBairro2 = rs3("nome_combo2")
 
  dim filename3,filename4,filename5
  filename3 = "imovel00000.jpg"
	 filename4 = "imovel00000.jpg"
  filename5 = "imovel00000.jpg"
	 filename6 = "imovel00000.jpg"
	  filename7 = "imovel00000.jpg"
	 
	 
	Conexao.execute"Insert into imoveis(proprietario,endereco,telefone,email,foto_grande,foto_pequena,foto_grande1,foto_grande2,foto_grande3,foto_grande4,foto_grande5,link_foto,cidade,bairro,tipo,area_total,area_construida,quartos,banheiros,vagas,negociacao,valor,data,obs_imovel,obs_proprietario,presenca_primeira,titulo_anuncio,texto_anuncio,standby,ocupacao) values( '"& vProprietario &"','"& vEndereco &"','"& vTelefone &"','"& vEmail &"','"& filename &"','"& filename2 &"','"& filename3 &"','"& filename4 &"','"& filename5 &"','"& filename6 &"','"& filename7 &"','"& vLink_Foto &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vArea_Total &"','"& vArea_Construida &"','"& vQuartos &"','"& vBanheiros &"','"& vVagas &"','"& vNegociacao &"','"& vValor &"','"& vdata &"','"& vObs_imovel &"','"& vObs_proprietario &"','"& vPresenca_primeira &"','"& vTitulo_anuncio &"','"& vTexto_anuncio &"','"& standby &"','"& vOcupacao &"')"
	'aqui os dados s�o colocados dentro do banco de dados.
	dim rs5,SQL5
 Set rs5 = Server.CreateObject("ADODB.RecordSet")
 SQL5 = "select * from imoveis ORDER BY cod_imovel DESC"
 rs5.open SQL5,Conexao,2,1
	 rs5.movefirst
	 dim varCodImovel5 
	 varCodImovel5 = rs5("cod_imovel")
	'aqui a vari�vel varCodImovel5 recebe o valor do c�digo do im�vel
	
	
	 
	
	  
      response.Redirect "form_incluir_imovel.asp?varSucesso_imovel="&vProprietario&""
      'voc� � redirecionado para o formul�rio de inclus�o de im�vel.







 
 
 
 
 
 
 
 
 
 
 
 
 
 
%>
<b>Uploaded file : </b><%=filename%><BR>
<img src="<%=filename%>"><br>

<!--#include file="upload.asp"-->
 
 <%
 end if
     
 
 
           
           Conexao.Close
		   rs2.close
		   rs3.close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>

