

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,varCodVisualiza_Paginas,rsFoto,vFoto,vNome,vTelefone,vEmail,strSQL2,vProposta,vimagem,vdata,SearchFor 
 varCodVisualiza_Paginas = request.QueryString("varCodVisualiza_Paginas")
 SearchFor = request.querystring("SearchFor")
 dim page
  page = request.querystring("page") 
 
 
 
   vimagem = "imovel00000.jpg"
   vdata = left(now(),8)
  vNome=request.form("txtNome")
      vTelefone=request.form("txtTelefone")
      vEmail=request.form("txtEmail")
	  vProposta=request.form("txtProposta")                                                           
															  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	
	
	
	Conexao.execute"Delete from visualiza_paginas where id_visualiza_pagina="&varCodVisualiza_Paginas 
	 
	 
	 
	 conexao.close
	 
	 set conexao = nothing
	 
	 
	 
	 response.redirect "archive_visualiza_paginas.asp?SearchFor="&SearchFor&"&page="&cInt(Page)&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">











 
 <%
 
     
 
 
           
          
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>


