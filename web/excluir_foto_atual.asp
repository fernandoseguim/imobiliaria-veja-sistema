

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<!--#include file="loggedin.asp"-->
<% response.buffer=True%>

<%
Dim Conexao,strSQL,rs,varCod_imovel,SearchFor,objFSO,FotoPeqEx,vProprietario,vEndereco 
dim rs3
dim strSQL3
dim varCidade, varBairro,varNegociacao,varQuartos,page,SearchWhere,varValor
dim varValor1,varValor2
dim varCodimovel
 varCodImovel = request.QueryString("varCodImovel")
                                                    
	
	dim varNomeFoto
	dim varNumFoto
	
	varNomeFoto = request.querystring("varNomeFoto")
	varNumFoto = request.querystring("varNumFoto")
	
	
															  
	Set rs = Server.CreateObject("ADODB.RecordSet")
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	
	
	
	Conexao.Open dsn
	
	



       
	
		
		
		If objFSO.FileExists(Server.MapPath(varNomeFoto)) = True  Then
	 objFSO.DeleteFile(Server.MapPath(varNomeFoto))
	 
	 end if
		
		if varNumFoto = "1" then
		
		Conexao.execute"update imoveis set foto_grande='"&"imovel00000.jpg"&"',foto_grande1='"&"imovel00000.jpg"&"',foto_pequena='"&"imovel00000.jpg"&"' where cod_imovel="&varCodImovel
		
	 
	 end if
	 
	 
	 if varNumFoto = "2" then
	 
	 Conexao.execute"update imoveis set foto_grande2='"&"imovel00000.jpg"&"' where cod_imovel="&varCodImovel
	 
	 end if
	 
	 
	 
	 if varNumFoto = "3" then
	 
	 Conexao.execute"update imoveis set foto_grande3='"&"imovel00000.jpg"&"' where cod_imovel="&varCodImovel
	 
	 end if
	 
	 
	 
	 if varNumFoto = "4" then
	 
	 Conexao.execute"update imoveis set foto_grande4='"&"imovel00000.jpg"&"' where cod_imovel="&varCodImovel
	 
	 end if
	 
	 
	 
	 if varNumFoto = "5" then
	 
	 Conexao.execute"update imoveis set foto_grande5='"&"imovel00000.jpg"&"' where cod_imovel="&varCodImovel
	 
	 end if
	 
	 
	 
	 if varNumFoto = "6" then
	 
	 Conexao.execute"update imoveis set foto_grande6='"&"imovel00000.jpg"&"' where cod_imovel="&varCodImovel
	 
	 end if
	 
	 
	 
	 
	 if varNumFoto = "7" then
	 
	 Conexao.execute"update imoveis set foto_grande7='"&"imovel00000.jpg"&"' where cod_imovel="&varCodImovel
	 
	 end if
	 
	 
	 
	 if varNumFoto = "8" then
	 
	 Conexao.execute"update imoveis set foto_grande8='"&"imovel00000.jpg"&"' where cod_imovel="&varCodImovel
	 
	 end if
	 
	 
	 
	 
	 if varNumFoto = "9" then
	 
	 Conexao.execute"update imoveis set foto_grande9='"&"imovel00000.jpg"&"' where cod_imovel="&varCodImovel
	 
	 end if
	 
	 
	 
	 
	 if varNumFoto = "10" then
	 
	 Conexao.execute"update imoveis set foto_grande10='"&"imovel00000.jpg"&"' where cod_imovel="&varCodImovel
	 
	 end if
	 
	response.redirect "visualizar_fotos.asp?varCodImovel="&varCodImovel&""
	
     
	  
	  
   
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">










 
 <%
 
     
 
 
           
           Conexao.Close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>



