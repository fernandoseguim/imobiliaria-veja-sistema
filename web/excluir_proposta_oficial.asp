

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,varCod_proposta_oficial,SearchFor
 varCod_proposta_oficial = request.QueryString("varCod_proposta_oficial")
SearchFor = request.Querystring("searchFor")
 
                                                            
															  
	Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	
	
	
	
	Conexao.Open dsn
	
	

        
	
	
	Conexao.execute"delete from proposta_oficial where cod_proposta_oficial="&VarCod_proposta_oficial 
	
	
	
	set rs = nothing
	
	conexao.close
	
	set conexao = nothing
	 
	 response.redirect "archive_proposta_oficial.asp?searchFor="&SearchFor&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
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




