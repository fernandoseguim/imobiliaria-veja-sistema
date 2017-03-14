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

if rs("foto_grande") <> "imovel00000.jpg" and rs("foto_grande1") <> "imovel00000.jpg"  and rs("foto_grande2") <> "imovel00000.jpg" and rs("foto_grande3") <> "imovel00000.jpg" and rs("foto_grande4") <> "imovel00000.jpg" and rs("foto_grande5") <> "imovel00000.jpg"  and rs("foto_grande6") <> "imovel00000.jpg"  and rs("foto_grande7") <> "imovel00000.jpg" and rs("foto_grande8") <> "imovel00000.jpg" and rs("foto_grande9") <> "imovel00000.jpg" and rs("foto_grande10") <> "imovel00000.jpg" then

response.Redirect "form_adicionar_foto.asp?varResultado="&"Não há espaço para mais fotos!!!"&"&varCodImovel="&varCodImovel&""
end if














contentType = UploadRequest.Item("txtAnuncio").Item("ContentType")
filepathname = UploadRequest.Item("txtAnuncio").Item("FileName")


filename = Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
filename = nome&filename




If objFSO.FileExists(Server.MapPath(filename)) = True then
response.redirect "form_adicionar_foto.asp?varResultado="&"Já existe uma foto com esse nome, tente outro!!!"&"&varCodImovel="&varCodImovel&""
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
   
  

if rs("foto_grande") = "imovel00000.jpg" or rs("foto_grande1") = "imovel00000.jpg"   then

	Conexao.execute"update imoveis set foto_grande='"&filename&"',foto_grande1='"&filename&"',foto_pequena='"&filename&"' where cod_imovel="&varCodImovel
  
  response.Redirect "form_adicionar_foto.asp?varResultado="&"Mais uma foto incluída!"&"&varCodImovel="&varCodImovel&""
   end if
   
   
   '-------------------------------------------------------------------------------
   
   
   
   
   if  rs("foto_grande2") = "imovel00000.jpg"  then

	Conexao.execute"update imoveis set foto_grande2='"&filename&"' where cod_imovel="&varCodImovel
  
  response.Redirect "form_adicionar_foto.asp?varResultado="&"Mais uma foto incluída!"&"&varCodImovel="&varCodImovel&""
   end if



'-----------------------------------------------------------------------------	 
	
	
	if rs("foto_grande3") = "imovel00000.jpg"  then

	Conexao.execute"update imoveis set foto_grande3='"&filename&"' where cod_imovel="&varCodImovel
  
   response.Redirect "form_adicionar_foto.asp?varResultado="&"Mais uma foto incluída!"&"&varCodImovel="&varCodImovel&""
   
   end if
	
	
	'------------------------------------------------------------------------------
	
	if rs("foto_grande4") = "imovel00000.jpg"  then

	Conexao.execute"update imoveis set foto_grande4='"&filename&"' where cod_imovel="&varCodImovel
  response.Redirect "form_adicionar_foto.asp?varResultado="&"Mais uma foto incluída!"&"&varCodImovel="&varCodImovel&""
  
   end if
	
	'------------------------------------------------------------------------------------------
	
	if rs("foto_grande5") = "imovel00000.jpg" then

	Conexao.execute"update imoveis set foto_grande5='"&filename&"' where cod_imovel="&varCodImovel
   response.Redirect "form_adicionar_foto.asp?varResultado="&"Mais uma foto incluída!"&"&varCodImovel="&varCodImovel&""
   
   end if
	
	'-----------------------------------------------------------------------------------------
	
	
	'------------------------------------------------------------------------------------------
	
	if rs("foto_grande6") = "imovel00000.jpg" then

	Conexao.execute"update imoveis set foto_grande6='"&filename&"' where cod_imovel="&varCodImovel
   response.Redirect "form_adicionar_foto.asp?varResultado="&"Mais uma foto incluída!"&"&varCodImovel="&varCodImovel&""
   
   end if
	
	'-----------------------------------------------------------------------------------------
	
	
	'------------------------------------------------------------------------------------------
	
	if rs("foto_grande7") = "imovel00000.jpg" then

	Conexao.execute"update imoveis set foto_grande7='"&filename&"' where cod_imovel="&varCodImovel
   response.Redirect "form_adicionar_foto.asp?varResultado="&"Mais uma foto incluída!"&"&varCodImovel="&varCodImovel&""
   
   end if
	
	'-----------------------------------------------------------------------------------------
	
	
	'------------------------------------------------------------------------------------------
	
	if rs("foto_grande8") = "imovel00000.jpg" then

	Conexao.execute"update imoveis set foto_grande8='"&filename&"' where cod_imovel="&varCodImovel
   response.Redirect "form_adicionar_foto.asp?varResultado="&"Mais uma foto incluída!"&"&varCodImovel="&varCodImovel&""
   
   end if
	
	'-----------------------------------------------------------------------------------------
	
	
	'------------------------------------------------------------------------------------------
	
	if rs("foto_grande9") = "imovel00000.jpg" then

	Conexao.execute"update imoveis set foto_grande9='"&filename&"' where cod_imovel="&varCodImovel
   response.Redirect "form_adicionar_foto.asp?varResultado="&"Mais uma foto incluída!"&"&varCodImovel="&varCodImovel&""
   
   end if
	
	'-----------------------------------------------------------------------------------------
	
	
	
	'------------------------------------------------------------------------------------------
	
	if rs("foto_grande10") = "imovel00000.jpg" then

	Conexao.execute"update imoveis set foto_grande10='"&filename&"' where cod_imovel="&varCodImovel
   response.Redirect "form_adicionar_foto.asp?varResultado="&"Mais uma foto incluída!"&"&varCodImovel="&varCodImovel&""
   
   end if
	
	'-----------------------------------------------------------------------------------------
	
	
	
	
	 
	 
     response.Redirect "form_adicionar_foto.asp?varResultado="&"Mais uma foto incluída!"&"&varCodImovel="&varCodImovel&""








 
 
 
 
 
 
 
 
 
 
 
 
 
 
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

