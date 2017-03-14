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
Dim UploadRequest
Set UploadRequest = CreateObject("Scripting.Dictionary")
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")


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

Conexao6.execute"update contador set hits='"&nome&"' where Cod_hits = 1"






BuildUploadRequest  RequestBin


Dim Conexao,rs,strSQL
 Dim varSucesso_promocao,vEmail,vTelefone,vTitulo,vData,vDescricao,vLiga_Desliga
 Dim varResultado

 Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	Set rs = Server.CreateObject("ADODB.RecordSet")
	
	dim varCodImovel
	varCodImovel = request.QueryString("varCodImovel")
strSQL = "SELECT * FROM imoveis where cod_imovel="&varCodImovel

rs.open strSQL,Conexao


if not rs.eof then



dim temFoto
	temFoto = rs("data_foto_colocada")
	
	
	if temFoto <> "" then
	
	response.write ""
	else 
	
	Conexao.execute"update imoveis set data_foto_colocada='"&now()&"',quem_foto_colocada='"&session("nome_id")&"' where cod_imovel="&varCodImovel
  
	end if
	












contentType = UploadRequest.Item("txtAnuncio").Item("ContentType")
filepathname = UploadRequest.Item("txtAnuncio").Item("FileName")


filename = Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
filename = nome&filename

If objFSO.FileExists(Server.MapPath(filename)) = True then
response.redirect "form_atualizar_foto.asp?varResultado="&"Já existe uma foto com esse nome, tente outro!!!"&"&varCodImovel="&varCodImovel&""
else



value = UploadRequest.Item("txtAnuncio").Item("Value")

'Create FileSytemObject Component
 Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")

'Create and Write to a File
 pathEnd = Len(Server.mappath(Request.ServerVariables("PATH_INFO")))-14
 Set MyFile = ScriptObject.CreateTextFile(Left(Server.mappath(Request.ServerVariables("PATH_INFO")),pathEnd)&filename)
 
 For i = 1 to LenB(value)
	 MyFile.Write chr(AscB(MidB(value,i,1)))
 Next
 
 MyFile.Close
 
 
 
 


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
   
 
 dim varNumFoto
 
 varNumFoto = request.querystring("varNumFoto")
 
 
  
 
 
  

if varNumFoto = "1"   then

If objFSO.FileExists(Server.MapPath(rs("Foto_grande1"))) = True  Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande1")))
	 
	 end if

	Conexao.execute"update imoveis set foto_grande='"&filename&"',foto_grande1='"&filename&"',foto_pequena='"&filename&"' where cod_imovel="&varCodImovel
  
  response.Redirect "form_atualizar_foto.asp?varResultado="&"Mais uma foto atualizada!"&"&varCodImovel="&varCodImovel&""
   end if
   
  
   
   if varNumFoto = "2"   then
   
   
   
		If objFSO.FileExists(Server.MapPath(rs("Foto_grande2"))) = True  Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande2")))
	 
	 end if
   

	Conexao.execute"update imoveis set foto_grande2='"&filename&"' where cod_imovel="&varCodImovel
  
  response.Redirect "form_atualizar_foto.asp?varResultado="&"Mais uma foto atualizada!"&"&varCodImovel="&varCodImovel&""
   end if 
  
   
   
   
   
   if varNumFoto = "3"   then
   
   
   
		If objFSO.FileExists(Server.MapPath(rs("Foto_grande3"))) = True  Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande3")))
	 
	 end if
   

	Conexao.execute"update imoveis set foto_grande3='"&filename&"' where cod_imovel="&varCodImovel
  
  response.Redirect "form_atualizar_foto.asp?varResultado="&"Mais uma foto atualizada!"&"&varCodImovel="&varCodImovel&""
   end if 
  
   
    
   if varNumFoto = "4"   then
   
   
   
		If objFSO.FileExists(Server.MapPath(rs("Foto_grande4"))) = True  Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande4")))
	 
	 end if
   

	Conexao.execute"update imoveis set foto_grande4='"&filename&"' where cod_imovel="&varCodImovel
  
  response.Redirect "form_atualizar_foto.asp?varResultado="&"Mais uma foto atualizada!"&"&varCodImovel="&varCodImovel&""
   end if 
  
   
   
   
    
   if varNumFoto = "5"   then
   
   
   
		If objFSO.FileExists(Server.MapPath(rs("Foto_grande5"))) = True  Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande5")))
	 
	 end if
   

	Conexao.execute"update imoveis set foto_grande5='"&filename&"' where cod_imovel="&varCodImovel
  
  response.Redirect "form_atualizar_foto.asp?varResultado="&"Mais uma foto atualizada!"&"&varCodImovel="&varCodImovel&""
   end if 
  
   
   
   
   
    
   if varNumFoto = "6"   then
   
   
   
		If objFSO.FileExists(Server.MapPath(rs("Foto_grande6"))) = True  Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande6")))
	 
	 end if
   

	Conexao.execute"update imoveis set foto_grande6='"&filename&"' where cod_imovel="&varCodImovel
  
  response.Redirect "form_atualizar_foto.asp?varResultado="&"Mais uma foto atualizada!"&"&varCodImovel="&varCodImovel&""
   end if 
  
   
   
   
   
   
    
   if varNumFoto = "7"   then
   
   
   
		If objFSO.FileExists(Server.MapPath(rs("Foto_grande7"))) = True  Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande7")))
	 
	 end if
   

	Conexao.execute"update imoveis set foto_grande7='"&filename&"' where cod_imovel="&varCodImovel
  
  response.Redirect "form_atualizar_foto.asp?varResultado="&"Mais uma foto atualizada!"&"&varCodImovel="&varCodImovel&""
   end if 
  
   
   
   
   
   
   
   
   
   
   
   
   
   if varNumFoto = "8"   then
   
   
   
		If objFSO.FileExists(Server.MapPath(rs("Foto_grande8"))) = True  Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande8")))
	 
	 end if
   

	Conexao.execute"update imoveis set foto_grande8='"&filename&"' where cod_imovel="&varCodImovel
  
  response.Redirect "form_atualizar_foto.asp?varResultado="&"Mais uma foto atualizada!"&"&varCodImovel="&varCodImovel&""
   end if 
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
    
   if varNumFoto = "9"   then
   
   
   
		If objFSO.FileExists(Server.MapPath(rs("Foto_grande9"))) = True  Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande9")))
	 
	 end if
   

	Conexao.execute"update imoveis set foto_grande9='"&filename&"' where cod_imovel="&varCodImovel
  
  response.Redirect "form_atualizar_foto.asp?varResultado="&"Mais uma foto atualizada!"&"&varCodImovel="&varCodImovel&""
   end if 
  
   
   
   
   
   
   
   
   
   
   
    
   if varNumFoto = "10"   then
   
   
   
		If objFSO.FileExists(Server.MapPath(rs("Foto_grande10"))) = True  Then
	 objFSO.DeleteFile(Server.MapPath(rs("Foto_grande10")))
	 
	 end if
   

	Conexao.execute"update imoveis set foto_grande10='"&filename&"' where cod_imovel="&varCodImovel
  
  response.Redirect "form_atualizar_foto.asp?varResultado="&"Mais uma foto atualizada!"&"&varCodImovel="&varCodImovel&""
   end if 
  
   
   
   '-------------------------------------------------------------------------------
   
   
   
  
	
	'-----------------------------------------------------------------------------------------
	
	 
	 
     response.Redirect "form_atualizar_foto.asp?varResultado="&"Mais uma foto incluída!"&"&varCodImovel="&varCodImovel&""








 
 
 
 
 
 
 
 
 
 
 
 
 
 
%>
<b>Uploaded file : </b><%=filename%><BR>
<img src="<%=filename%>"><br>

<!--#include file="upload.asp"-->
 
 <%
 end if
     end if
 
 
           rs.close
           Conexao.Close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>

