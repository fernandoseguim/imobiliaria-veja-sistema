

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,varCodProposta,rsFoto,vFoto,vNome,vTelefone,vEmail,strSQL2,vProposta,vimagem,vdata,SearchFor 
 varCodProposta = request.QueryString("varCodProposta")

 dim varCod_telefone_acesso
 
 SearchFor = request.querystring("SearchFor")
 
 varCod_telefone_acesso = request.querystring("varCod_telefone_acesso")
 
 
 
                                                             
															  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	
	
	
	Conexao.execute"Delete from telefone_acesso where cod_telefone_acesso="&VarCod_telefone_acesso 
	 
	 response.redirect "archive_telefone_acesso.asp?SearchFor="&SearchFor&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
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


